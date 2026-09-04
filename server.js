// const express = require("express");
// const cors = require("cors");
// const session = require("express-session");
// const msal = require("@azure/msal-node");
// const crypto = require("crypto");
// require("dotenv").config();

const express = require("express");
const cors = require("cors");
const session = require("express-session");
const msal = require("@azure/msal-node");
const crypto = require("crypto");
const OpenAI = require("openai");
require("dotenv").config();

const app = express();

// =====================================================
// PRISM AI LLM - HUGGING FACE
// =====================================================

const llmClient = new OpenAI({
  baseURL: "https://router.huggingface.co/v1",
  apiKey: process.env.HF_TOKEN,
});

const LLM_MODEL = "Qwen/Qwen3-4B-Instruct-2507";

// =====================================================
// PRISM AI CONNECTION STORE
// =====================================================

// Each Power BI custom visual gets its own connection ID.
//
// connectionId -> {
//     authenticated,
//     user,
//     accessToken
// }

const powerBiConnections = new Map();

// =====================================================
// BASIC EXPRESS CONFIGURATION
// =====================================================

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }

  next();
});

app.use(express.json());

// app.options("*", cors());

app.use(
  cors({
    origin: true,
    credentials: true,
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  }),
);

app.use(express.json());

// =====================================================
// SESSION
// =====================================================

app.use(
  session({
    secret: process.env.SESSION_SECRET || "bi-assistant-hackathon-secret",

    resave: false,

    saveUninitialized: false,

    cookie: {
      secure: false,
      httpOnly: true,
      sameSite: "lax",
    },
  }),
);

// =====================================================
// MICROSOFT ENTRA CONFIGURATION
// =====================================================

// IMPORTANT:
// Put Vinod's Tenant ID in .env.
//
// Example:
//
// TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
//
// We use the tenant directly instead of "common"
// because this hackathon app is being used inside
// Vinod's Microsoft Entra tenant.

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,

    // authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,

    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,

    clientSecret: process.env.CLIENT_SECRET,
  },
};

console.log("=================================");
console.log("MSAL CONFIGURATION");
console.log("Client ID:", process.env.CLIENT_ID);
console.log("Tenant ID:", process.env.TENANT_ID);
console.log("Redirect URI:", process.env.REDIRECT_URI);
console.log("=================================");

const cca = new msal.ConfidentialClientApplication(msalConfig);

const POWER_BI_SCOPES = [
  "https://analysis.windows.net/powerbi/api/Workspace.Read.All",
  "https://analysis.windows.net/powerbi/api/Report.Read.All",
  "https://analysis.windows.net/powerbi/api/Dataset.Read.All",
];

// const FABRIC_SCOPES = [
//   "https://api.fabric.microsoft.com/SemanticModel.ReadWrite.All",
// ];

// =====================================================
// POWER BI API
// =====================================================

const POWER_BI_API = "https://api.powerbi.com/v1.0/myorg";

// =====================================================
// POWER BI RESOURCE
// =====================================================

const POWER_BI_RESOURCE = "https://analysis.windows.net/powerbi/api";

// .default means:
// "Give me the delegated permissions that this
// application has configured/consented for Power BI."

const POWER_BI_SCOPE = `${POWER_BI_RESOURCE}/.default`;

// =====================================================
// HOME
// =====================================================

app.get("/", (req, res) => {
  res.send("BI Assistant backend is running successfully.");
});

// =====================================================
// TEST
// =====================================================

app.get("/test", (req, res) => {
  res.json({
    success: true,
    message: "TEST ROUTE WORKS",
  });
});

// =====================================================
// CREATE PRISM AI CONNECTION
// =====================================================

app.get("/api/create-connection", (req, res) => {
  const connectionId = crypto.randomUUID();

  powerBiConnections.set(connectionId, {
    authenticated: false,
    user: null,
    accessToken: null,
  });

  console.log("=================================");
  console.log("NEW PRISM AI CONNECTION CREATED");
  console.log("Connection ID:", connectionId);
  console.log("=================================");

  res.json({
    connectionId,
  });
});

// =====================================================
// LOGIN
// =====================================================

app.get("/login", async (req, res) => {
  try {
    const connectionId = req.query.connectionId;

    if (!connectionId) {
      return res.status(400).send("Missing PRISM AI connection ID.");
    }

    const connection = powerBiConnections.get(connectionId);

    if (!connection) {
      return res.status(400).send("Invalid or expired PRISM AI connection.");
    }

    console.log("=================================");
    console.log("STARTING MICROSOFT LOGIN");
    console.log("Connection ID:", connectionId);
    console.log("=================================");

    const authCodeUrlParameters = {
      scopes: POWER_BI_SCOPES,

      redirectUri: process.env.REDIRECT_URI,

      prompt: "select_account",

      // Microsoft returns this value
      // back to /auth/callback
      state: connectionId,
    };

    const response = await cca.getAuthCodeUrl(authCodeUrlParameters);

    res.redirect(response);
  } catch (error) {
    console.error("Login error:", error);

    res.status(500).send("Login failed.");
  }
});

// =====================================================
// AUTHENTICATION CALLBACK
// =====================================================

app.get("/auth/callback", async (req, res) => {
  console.log("=================================");
  console.log("AUTH CALLBACK RECEIVED");
  console.log("Query parameters:", req.query);
  console.log("Authorization code exists:", !!req.query.code);
  console.log("=================================");

  // =====================================
  // MICROSOFT AUTHENTICATION ERROR
  // =====================================

  if (req.query.error) {
    console.error("Microsoft OAuth error:", req.query.error);

    console.error("Description:", req.query.error_description);

    return res.status(400).send(`
      <h2>Microsoft Authentication Error</h2>

      <p>
        <b>Error:</b>
        ${req.query.error}
      </p>

      <p>
        <b>Description:</b>
        ${req.query.error_description || "No description"}
      </p>
    `);
  }

  // =====================================
  // NO AUTHORIZATION CODE
  // =====================================

  if (!req.query.code) {
    return res.status(400).send(`
      <h2>Authentication Error</h2>

      <p>
        No authorization code was received from Microsoft.
      </p>

      <p>
        Please start again from the BI Assistant.
      </p>
    `);
  }

  try {
    // =====================================
    // GET CONNECTION ID
    // =====================================

    const connectionId = req.query.state;

    console.log("Connection ID from OAuth state:", connectionId);

    if (!connectionId) {
      return res.status(400).send("Authentication connection ID is missing.");
    }

    // =====================================
    // FIND PRISM AI CONNECTION
    // =====================================

    const connection = powerBiConnections.get(connectionId);

    if (!connection) {
      return res.status(400).send("Invalid or expired PRISM AI connection.");
    }

    // =====================================
    // EXCHANGE AUTHORIZATION CODE FOR TOKEN
    // =====================================

    const tokenRequest = {
      code: req.query.code,

      scopes: POWER_BI_SCOPES,

      redirectUri: process.env.REDIRECT_URI,
    };

    const response = await cca.acquireTokenByCode(tokenRequest);

    console.log("=================================");
    console.log("USER AUTHENTICATED SUCCESSFULLY");

    console.log("Username:", response.account?.username);

    console.log("Token obtained:", !!response.accessToken);

    console.log("Expires:", response.expiresOn);

    console.log("=================================");

    // =====================================
    // SAVE AUTHENTICATION TO CONNECTION
    // =====================================

    connection.authenticated = true;

    connection.user = {
      username: response.account?.username,

      name: response.account?.name,
    };

    connection.accessToken = response.accessToken;

    // Update connection store

    powerBiConnections.set(connectionId, connection);

    console.log("=================================");
    console.log("PRISM AI CONNECTION UPDATED");

    console.log("Connection ID:", connectionId);

    console.log("Authenticated:", connection.authenticated);

    console.log("User:", connection.user?.username);

    console.log("Access Token Exists:", !!connection.accessToken);

    console.log("=================================");

    // =====================================
    // SUCCESS PAGE
    // =====================================

    res.send(`
      <!DOCTYPE html>

      <html>

      <head>

        <title>
          Power BI Connected
        </title>

      </head>


      <body style="
        font-family: Arial;
        text-align: center;
        padding-top: 100px;
      ">

        <h2>
          ✅ Power BI Connected Successfully!
        </h2>


        <p>
          You can now return to PRISM AI.
        </p>


        <p>
          This window will close automatically.
        </p>


        <script>

          setTimeout(() => {

            window.close();

          }, 2000);

        </script>

      </body>

      </html>
    `);
  } catch (error) {
    console.error("=================================");
    console.error("TOKEN ACQUISITION ERROR");

    console.error(error);

    console.log("=================================");

    res.status(500).send(`

      <h2>
        Authentication failed
      </h2>

      <pre>
        ${error.message}
      </pre>

    `);
  }
});

// =====================================================
// DEBUG POWER BI TOKEN
// =====================================================

app.get("/api/debug-token", (req, res) => {
  try {
    // =================================================
    // CHECK TOKEN
    // =================================================

    const token = req.session.powerBiAccessToken;

    if (!token) {
      return res.status(401).json({
        authenticated: false,

        error: "No Power BI access token found in session.",
      });
    }

    // =================================================
    // JWT STRUCTURE
    // =================================================

    const parts = token.split(".");

    if (parts.length !== 3) {
      return res.status(400).json({
        error: "Stored token is not a JWT.",
      });
    }

    // =================================================
    // DECODE PAYLOAD
    // =================================================

    const payload = JSON.parse(
      Buffer.from(parts[1], "base64url").toString("utf8"),
    );

    // =================================================
    // RETURN SAFE TOKEN INFORMATION
    // =================================================

    res.json({
      authenticated: true,

      username: req.session.user?.username,

      tokenInfo: {
        audience: payload.aud,

        issuer: payload.iss,

        tenantId: payload.tid,

        username: payload.preferred_username,

        scopes: payload.scp,

        expiresAt: payload.exp
          ? new Date(payload.exp * 1000).toISOString()
          : null,

        issuedAt: payload.iat
          ? new Date(payload.iat * 1000).toISOString()
          : null,
      },
    });
  } catch (error) {
    console.error("Debug token error:", error);

    res.status(500).json({
      error: "Unable to decode token.",

      details: error.message,
    });
  }
});

// =====================================================
// TEST POWER BI API
// =====================================================

app.get("/api/test-powerbi", async (req, res) => {
  try {
    // =================================================
    // CHECK AUTHENTICATION
    // =================================================

    // if (!req.session.powerBiAccessToken) {
    //   return res.status(401).json({
    //     authenticated: false,

    //     error: "User is not authenticated.",
    //   });
    // }

    console.log("========== PAGES DEBUG ==========");
    console.log("Session ID:", req.sessionID);
    console.log("Session:", req.session);
    console.log(
      "Power BI Access Token exists:",
      !!req.session.powerBiAccessToken,
    );
    console.log("=================================");

    if (!req.session.powerBiAccessToken) {
      return res.status(401).json({
        error: "Not authenticated",
        sessionId: req.sessionID,
        tokenExists: false,
      });
    }

    const accessToken = req.session.powerBiAccessToken;

    console.log("Testing Power BI API...");

    // =================================================
    // CALL SIMPLE POWER BI API
    // =================================================

    const response = await fetch(`${POWER_BI_API}/groups?$top=10`, {
      method: "GET",

      headers: {
        Authorization: `Bearer ${accessToken}`,

        Accept: "application/json",
      },
    });

    // =================================================
    // READ RESPONSE
    // =================================================

    const responseText = await response.text();

    console.log("Power BI status:", response.status);

    console.log("Power BI response:", responseText);

    // =================================================
    // POWER BI ERROR
    // =================================================

    if (!response.ok) {
      return res.status(response.status).json({
        success: false,

        powerBiStatus: response.status,

        powerBiStatusText: response.statusText,

        powerBiResponse: responseText,
      });
    }

    // =================================================
    // SUCCESS
    // =================================================

    let data;

    try {
      data = JSON.parse(responseText);
    } catch {
      data = responseText;
    }

    res.json({
      success: true,

      message: "Power BI API is working.",

      data,
    });
  } catch (error) {
    console.error("Power BI test error:", error);

    res.status(500).json({
      success: false,

      error: error.message,
    });
  }
});

// =====================================================
// GET POWER BI WORKSPACES
// =====================================================

// =====================================================
// GET POWER BI WORKSPACES
// =====================================================

app.get("/api/workspaces", async (req, res) => {
  const connectionId = req.query.connectionId;

  console.log("=================================");
  console.log("GETTING POWER BI WORKSPACES");
  console.log("Connection ID:", connectionId);
  console.log("=================================");

  // Check connection ID

  if (!connectionId) {
    return res.status(400).json({
      error: "Connection ID is missing.",
    });
  }

  // Find PRISM AI connection

  const connection = powerBiConnections.get(connectionId);

  if (!connection) {
    return res.status(401).json({
      error: "Invalid PRISM AI connection.",
    });
  }

  // Check authentication

  if (!connection.authenticated || !connection.accessToken) {
    return res.status(401).json({
      error: "Power BI authentication required.",
    });
  }

  try {
    const response = await fetch("https://api.powerbi.com/v1.0/myorg/groups", {
      method: "GET",

      headers: {
        Authorization: `Bearer ${connection.accessToken}`,

        "Content-Type": "application/json",
      },
    });

    const data = await response.json();

    console.log("POWER BI WORKSPACES RESPONSE:", data);

    if (!response.ok) {
      return res.status(response.status).json(data);
    }

    res.json(data);
  } catch (error) {
    console.error("WORKSPACE ERROR:", error);

    res.status(500).json({
      error: "Failed to retrieve Power BI workspaces.",
    });
  }
});

// =====================================================
// GET ONE SPECIFIC WORKSPACE
// =====================================================

app.get("/api/workspace/:workspaceId", async (req, res) => {
  try {
    if (!req.session.powerBiAccessToken) {
      return res.status(401).json({
        error: "User is not authenticated.",
      });
    }

    const workspaceId = req.params.workspaceId;

    const accessToken = req.session.powerBiAccessToken;

    console.log("Testing workspace:", workspaceId);

    const response = await fetch(`${POWER_BI_API}/groups/${workspaceId}`, {
      method: "GET",

      headers: {
        Authorization: `Bearer ${accessToken}`,

        Accept: "application/json",
      },
    });

    const responseText = await response.text();

    if (!response.ok) {
      return res.status(response.status).json({
        success: false,

        status: response.status,

        statusText: response.statusText,

        details: responseText,
      });
    }

    res.json({
      success: true,

      workspace: JSON.parse(responseText),
    });
  } catch (error) {
    console.error("Workspace test error:", error);

    res.status(500).json({
      error: error.message,
    });
  }
});

// =====================================================
// START SERVER
// =====================================================

// // =====================================================
// // GET DATASETS / SEMANTIC MODELS IN A WORKSPACE
// // =====================================================

const PORT = process.env.PORT || 3000;

app.get("/api/workspaces/:workspaceId/reports", async (req, res) => {
  try {
    if (!req.session.powerBiAccessToken) {
      return res.status(401).json({
        error: "User is not authenticated.",
      });
    }

    const workspaceId = req.params.workspaceId;
    const accessToken = req.session.powerBiAccessToken;

    console.log("=================================");
    console.log("GETTING REPORTS");
    console.log("Workspace ID:", workspaceId);
    console.log("=================================");

    const response = await fetch(
      `${POWER_BI_API}/groups/${workspaceId}/reports`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
        },
      },
    );

    const text = await response.text();

    if (!response.ok) {
      console.error("Power BI Reports API error:");
      console.error("Status:", response.status);
      console.error("Response:", text);

      return res.status(response.status).json({
        error: "Unable to retrieve reports.",
        status: response.status,
        details: text,
      });
    }

    const data = JSON.parse(text);

    console.log("Reports retrieved:");
    console.log(data.value);

    res.json(data);
  } catch (error) {
    console.error("Reports retrieval error:", error);

    res.status(500).json({
      error: "Failed to retrieve reports.",
      details: error.message,
    });
  }
});

// ========================================
// GET REPORT PAGES
// ========================================

// ========================================
// GET REPORT PAGES
// ========================================

app.get(
  "/api/workspaces/:workspaceId/reports/:reportId/pages",
  async (req, res) => {
    try {
      const { workspaceId, reportId } = req.params;

      // Check authentication
      if (!req.session.powerBiAccessToken) {
        return res.status(401).json({
          error: "User is not authenticated.",
        });
      }

      const accessToken = req.session.powerBiAccessToken;

      console.log("=================================");
      console.log("GETTING REPORT PAGES");
      console.log("Workspace ID:", workspaceId);
      console.log("Report ID:", reportId);
      console.log("Token exists:", !!accessToken);
      console.log("=================================");

      const response = await fetch(
        `${POWER_BI_API}/groups/${workspaceId}/reports/${reportId}/pages`,
        {
          method: "GET",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json",
          },
        },
      );

      const text = await response.text();

      console.log("Power BI Pages API status:", response.status);

      if (!response.ok) {
        console.error("Power BI Pages API error:");
        console.error(text);

        return res.status(response.status).json({
          error: "Unable to retrieve report pages.",
          status: response.status,
          details: text,
        });
      }

      const data = JSON.parse(text);

      console.log("Report pages retrieved:");
      console.log(data);

      res.json(data);
    } catch (error) {
      console.error("Report pages retrieval error:", error);

      res.status(500).json({
        error: "Failed to retrieve report pages.",
        details: error.message,
      });
    }
  },
);

// =====================================================
// GET DATASETS / SEMANTIC MODELS IN A WORKSPACE
// =====================================================

// =====================================================
// GET SEMANTIC MODELS / DATASETS
// =====================================================

app.get(
  "/api/workspaces/:workspaceId/datasets",

  async (req, res) => {
    const workspaceId = req.params.workspaceId;

    const connectionId = req.query.connectionId;

    console.log("=================================");
    console.log("GETTING SEMANTIC MODELS");
    console.log("Workspace ID:", workspaceId);
    console.log("Connection ID:", connectionId);
    console.log("=================================");

    // Check connection ID

    if (!connectionId) {
      return res.status(400).json({
        error: "Connection ID is missing.",
      });
    }

    // Find PRISM AI connection

    const connection = powerBiConnections.get(connectionId);

    if (!connection) {
      return res.status(401).json({
        error: "Invalid PRISM AI connection.",
      });
    }

    // Check authentication

    if (!connection.authenticated || !connection.accessToken) {
      return res.status(401).json({
        error: "Power BI authentication required.",
      });
    }

    try {
      const response = await fetch(
        `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/datasets`,

        {
          method: "GET",

          headers: {
            Authorization: `Bearer ${connection.accessToken}`,

            "Content-Type": "application/json",
          },
        },
      );

      const data = await response.json();

      console.log("SEMANTIC MODELS RESPONSE:", data);

      if (!response.ok) {
        return res.status(response.status).json(data);
      }

      res.json(data);
    } catch (error) {
      console.error("SEMANTIC MODEL ERROR:", error);

      res.status(500).json({
        error: "Failed to retrieve semantic models.",
      });
    }
  },
);

// =====================================================
// AUTH STATUS
// =====================================================

app.get("/api/auth-status", (req, res) => {
  const connectionId = req.query.connectionId;

  console.log("=================================");
  console.log("AUTH STATUS CHECK");
  console.log("Connection ID:", connectionId);
  console.log("=================================");

  if (!connectionId) {
    return res.status(400).json({
      authenticated: false,
      error: "Connection ID is missing.",
    });
  }

  const connection = powerBiConnections.get(connectionId);

  if (!connection) {
    console.log("CONNECTION NOT FOUND");

    return res.json({
      authenticated: false,
      user: null,
    });
  }

  console.log("Authenticated:", connection.authenticated);

  console.log("User:", connection.user);

  console.log("=================================");

  res.json({
    authenticated: connection.authenticated === true,

    user: connection.user || null,
  });
});

// =====================================================
// TEST DAX QUERY
// =====================================================

app.get(
  "/api/workspaces/:workspaceId/datasets/:datasetId/test-query",
  async (req, res) => {
    try {
      if (!req.session.powerBiAccessToken) {
        return res.status(401).json({
          error: "User is not authenticated.",
        });
      }

      const { workspaceId, datasetId } = req.params;

      const accessToken = req.session.powerBiAccessToken;

      // -------------------------------------------------
      // TEMPORARY TEST DAX
      // -------------------------------------------------

      const daxQuery = `
        EVALUATE
        ROW(
          "Test", 1
        )
      `;

      console.log("=================================");
      console.log("EXECUTING DAX QUERY");
      console.log("Workspace:", workspaceId);
      console.log("Dataset:", datasetId);
      console.log("DAX:", daxQuery);
      console.log("=================================");

      const response = await fetch(
        `${POWER_BI_API}/groups/${workspaceId}/datasets/${datasetId}/executeQueries`,
        {
          method: "POST",

          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },

          body: JSON.stringify({
            queries: [
              {
                query: daxQuery,
              },
            ],

            serializerSettings: {
              includeNulls: true,
            },
          }),
        },
      );

      const responseText = await response.text();

      console.log("Execute Queries API Status:", response.status);

      console.log("Execute Queries API Response:", responseText);

      if (!response.ok) {
        return res.status(response.status).json({
          error: "Unable to execute DAX query.",
          status: response.status,
          details: responseText,
        });
      }

      const data = JSON.parse(responseText);

      res.json({
        success: true,
        daxQuery,
        result: data,
      });
    } catch (error) {
      console.error("DAX QUERY ERROR:", error);

      res.status(500).json({
        error: "Failed to execute DAX query.",
        details: error.message,
      });
    }
  },
);

app.get("/api/health", (req, res) => {
  res.json({
    success: true,
    message: "Backend reachable from Power BI visual",
    sessionId: req.sessionID,
    authenticated: !!req.session.powerBiAccessToken,
  });
});

// =====================================================
// GET SEMANTIC MODEL METADATA
// =====================================================

app.get(
  "/api/workspaces/:workspaceId/datasets/:datasetId/metadata",

  async (req, res) => {
    const { workspaceId, datasetId } = req.params;

    const connectionId = req.query.connectionId;

    console.log("=================================");
    console.log("GETTING SEMANTIC MODEL METADATA");
    console.log("Workspace ID:", workspaceId);
    console.log("Dataset ID:", datasetId);
    console.log("Connection ID:", connectionId);
    console.log("=================================");

    if (!connectionId) {
      return res.status(400).json({
        error: "Connection ID is missing.",
      });
    }

    const connection = powerBiConnections.get(connectionId);

    if (!connection || !connection.authenticated || !connection.accessToken) {
      return res.status(401).json({
        error: "Power BI authentication required.",
      });
    }

    try {
      /*
       * First, get basic dataset information.
       */

      const datasetResponse = await fetch(
        `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/datasets/${datasetId}`,

        {
          headers: {
            Authorization: `Bearer ${connection.accessToken}`,
          },
        },
      );

      const datasetData = await datasetResponse.json();

      if (!datasetResponse.ok) {
        console.error("DATASET METADATA ERROR:", datasetData);

        return res.status(datasetResponse.status).json(datasetData);
      }

      console.log("SEMANTIC MODEL INFO:", datasetData);

      res.json({
        dataset: datasetData,

        message: "Basic semantic model information retrieved successfully.",
      });
    } catch (error) {
      console.error("SEMANTIC MODEL METADATA ERROR:", error);

      res.status(500).json({
        error: "Failed to retrieve semantic model metadata.",
      });
    }
  },
);

// =====================================================
// READ-ONLY SEMANTIC MODEL SCHEMA TEST
// =====================================================

app.get(
  "/api/workspaces/:workspaceId/datasets/:datasetId/schema",
  async (req, res) => {
    const { workspaceId, datasetId } = req.params;

    const connectionId = req.query.connectionId;

    console.log("=================================");
    console.log("PRISM AI - SCHEMA TEST");
    console.log("Workspace ID:", workspaceId);
    console.log("Dataset ID:", datasetId);
    console.log("Connection ID:", connectionId);
    console.log("=================================");

    // Check connection ID
    if (!connectionId) {
      return res.status(400).json({
        success: false,
        error: "Connection ID is missing.",
      });
    }

    // Get connection
    const connection = powerBiConnections.get(connectionId);

    // Check authentication
    if (!connection || !connection.authenticated || !connection.accessToken) {
      return res.status(401).json({
        success: false,
        error: "Power BI authentication required.",
      });
    }

    try {
      // -----------------------------------------
      // READ-ONLY TEST QUERY
      // -----------------------------------------

      const daxQuery = `
        EVALUATE
        ROW(
          "PRISM_AI_SCHEMA_TEST",
          1
        )
      `;

      console.log("Executing DAX:");
      console.log(daxQuery);

      // -----------------------------------------
      // EXECUTE AGAINST SELECTED DATASET
      // -----------------------------------------

      const response = await fetch(
        `${POWER_BI_API}/groups/${workspaceId}/datasets/${datasetId}/executeQueries`,
        {
          method: "POST",

          headers: {
            Authorization: `Bearer ${connection.accessToken}`,

            "Content-Type": "application/json",
          },

          body: JSON.stringify({
            queries: [
              {
                query: daxQuery,
              },
            ],

            serializerSettings: {
              includeNulls: true,
            },
          }),
        },
      );

      const responseText = await response.text();

      console.log("Power BI status:", response.status);

      console.log("Power BI response:", responseText);

      // -----------------------------------------
      // HANDLE POWER BI ERROR
      // -----------------------------------------

      if (!response.ok) {
        return res.status(response.status).json({
          success: false,

          error: "Unable to query the selected semantic model.",

          powerBiStatus: response.status,

          details: responseText,
        });
      }

      // -----------------------------------------
      // PARSE RESULT
      // -----------------------------------------

      const result = JSON.parse(responseText);

      // -----------------------------------------
      // SUCCESS
      // -----------------------------------------

      res.json({
        success: true,

        message: "PRISM AI can read the selected semantic model.",

        workspaceId,

        datasetId,

        readOnly: true,

        result,
      });
    } catch (error) {
      console.error("SCHEMA TEST ERROR:", error);

      res.status(500).json({
        success: false,

        error: "Semantic model schema test failed.",

        details: error.message,
      });
    }
  },
);

// =====================================================
// TEST SEMANTIC MODEL SCHEMA DISCOVERY
// =====================================================

app.get(
  "/api/workspaces/:workspaceId/datasets/:datasetId/schema-test",

  async (req, res) => {
    const { workspaceId, datasetId } = req.params;

    const connectionId = req.query.connectionId;

    console.log("=================================");
    console.log("SEMANTIC MODEL SCHEMA TEST");
    console.log("Workspace ID:", workspaceId);
    console.log("Dataset ID:", datasetId);
    console.log("Connection ID:", connectionId);
    console.log("=================================");

    if (!connectionId) {
      return res.status(400).json({
        error: "Connection ID is missing.",
      });
    }

    const connection = powerBiConnections.get(connectionId);

    if (!connection || !connection.authenticated || !connection.accessToken) {
      return res.status(401).json({
        error: "Power BI authentication required.",
      });
    }

    try {
      const daxQuery = `
        EVALUATE
        ROW(
          "PRISM_AI_Test", 1
        )
      `;

      const response = await fetch(
        `${POWER_BI_API}/groups/${workspaceId}/datasets/${datasetId}/executeQueries`,

        {
          method: "POST",

          headers: {
            Authorization: `Bearer ${connection.accessToken}`,

            "Content-Type": "application/json",
          },

          body: JSON.stringify({
            queries: [
              {
                query: daxQuery,
              },
            ],

            serializerSettings: {
              includeNulls: true,
            },
          }),
        },
      );

      const responseText = await response.text();

      console.log("SCHEMA TEST STATUS:", response.status);

      console.log("SCHEMA TEST RESPONSE:", responseText);

      if (!response.ok) {
        return res.status(response.status).json({
          error: "Unable to execute query against semantic model.",

          details: responseText,
        });
      }

      const data = JSON.parse(responseText);

      res.json({
        success: true,

        message:
          "PRISM AI successfully executed a query against the semantic model.",

        result: data,
      });
    } catch (error) {
      console.error("SCHEMA TEST ERROR:", error);

      res.status(500).json({
        error: "Schema test failed.",

        details: error.message,
      });
    }
  },
);

// =====================================================
// KPI - TOTAL SALES
// READ-ONLY QUERY
// =====================================================

app.get(
  "/api/workspaces/:workspaceId/datasets/:datasetId/kpi/total-sales",
  async (req, res) => {
    const { workspaceId, datasetId } = req.params;
    const connectionId = req.query.connectionId;

    console.log("=================================");
    console.log("KPI: TOTAL SALES");
    console.log("Workspace ID:", workspaceId);
    console.log("Dataset ID:", datasetId);
    console.log("Connection ID:", connectionId);
    console.log("=================================");

    // -----------------------------------------
    // CHECK CONNECTION ID
    // -----------------------------------------

    if (!connectionId) {
      return res.status(400).json({
        success: false,
        error: "Connection ID is missing.",
      });
    }

    // -----------------------------------------
    // GET CONNECTION
    // -----------------------------------------

    const connection = powerBiConnections.get(connectionId);

    if (!connection) {
      return res.status(401).json({
        success: false,
        error: "Invalid or expired PRISM AI connection.",
      });
    }

    // -----------------------------------------
    // CHECK AUTHENTICATION
    // -----------------------------------------

    if (!connection.authenticated || !connection.accessToken) {
      return res.status(401).json({
        success: false,
        error: "Power BI authentication required.",
      });
    }

    try {
      // -----------------------------------------
      // READ-ONLY DAX QUERY
      // -----------------------------------------
      //
      // IMPORTANT:
      // Replace Orders[Sales] below with the
      // actual table and column from your dataset.
      //
      // -----------------------------------------

      const daxQuery = `
        EVALUATE
        ROW(
          "TotalSales",
          SUM('Orders'[Sales])
        )
      `;

      console.log("Executing Total Sales DAX:");
      console.log(daxQuery);

      // -----------------------------------------
      // EXECUTE QUERY AGAINST POWER BI
      // -----------------------------------------

      const response = await fetch(
        `${POWER_BI_API}/groups/${workspaceId}/datasets/${datasetId}/executeQueries`,
        {
          method: "POST",

          headers: {
            Authorization: `Bearer ${connection.accessToken}`,
            "Content-Type": "application/json",
          },

          body: JSON.stringify({
            queries: [
              {
                query: daxQuery,
              },
            ],

            serializerSettings: {
              includeNulls: true,
            },
          }),
        },
      );

      const responseText = await response.text();

      console.log("Total Sales API status:", response.status);
      console.log("Total Sales response:", responseText);

      // -----------------------------------------
      // POWER BI ERROR
      // -----------------------------------------

      if (!response.ok) {
        return res.status(response.status).json({
          success: false,
          error: "Unable to calculate Total Sales.",
          powerBiStatus: response.status,
          details: responseText,
        });
      }

      // -----------------------------------------
      // PARSE POWER BI RESULT
      // -----------------------------------------

      const data = JSON.parse(responseText);

      // -----------------------------------------
      // EXTRACT TOTAL SALES
      // -----------------------------------------

      const totalSales =
        data?.results?.[0]?.tables?.[0]?.rows?.[0]?.["[TotalSales]"];

      console.log("TOTAL SALES VALUE:", totalSales);
      console.log("FULL POWER BI RESULT:", JSON.stringify(data, null, 2));

      // -----------------------------------------
      // RETURN RESULT
      // -----------------------------------------

      res.json({
        success: true,
        kpi: "Total Sales",
        value: totalSales,
        readOnly: true,
        daxQuery,
      });
    } catch (error) {
      console.error("TOTAL SALES ERROR:", error);

      res.status(500).json({
        success: false,
        error: "Failed to retrieve Total Sales.",
        details: error.message,
      });
    }
  },
);

// =====================================================
// KPI - TOTAL PROFIT
// READ-ONLY QUERY
// =====================================================

app.get(
  "/api/workspaces/:workspaceId/datasets/:datasetId/kpi/total-profit",
  async (req, res) => {
    const { workspaceId, datasetId } = req.params;
    const connectionId = req.query.connectionId;

    console.log("=================================");
    console.log("KPI: TOTAL PROFIT");
    console.log("Workspace ID:", workspaceId);
    console.log("Dataset ID:", datasetId);
    console.log("Connection ID:", connectionId);
    console.log("=================================");

    if (!connectionId) {
      return res.status(400).json({
        success: false,
        error: "Connection ID is missing.",
      });
    }

    const connection = powerBiConnections.get(connectionId);

    if (!connection) {
      return res.status(401).json({
        success: false,
        error: "Invalid or expired PRISM AI connection.",
      });
    }

    if (!connection.authenticated || !connection.accessToken) {
      return res.status(401).json({
        success: false,
        error: "Power BI authentication required.",
      });
    }

    try {
      const daxQuery = `
        EVALUATE
        ROW(
          "TotalProfit",
          SUM('Orders'[Profit])
        )
      `;

      console.log("Executing Total Profit DAX:");
      console.log(daxQuery);

      const response = await fetch(
        `${POWER_BI_API}/groups/${workspaceId}/datasets/${datasetId}/executeQueries`,
        {
          method: "POST",

          headers: {
            Authorization: `Bearer ${connection.accessToken}`,
            "Content-Type": "application/json",
          },

          body: JSON.stringify({
            queries: [
              {
                query: daxQuery,
              },
            ],

            serializerSettings: {
              includeNulls: true,
            },
          }),
        },
      );

      const responseText = await response.text();

      console.log("Total Profit API status:", response.status);
      console.log("Total Profit response:", responseText);

      if (!response.ok) {
        return res.status(response.status).json({
          success: false,
          error: "Unable to calculate Total Profit.",
          powerBiStatus: response.status,
          details: responseText,
        });
      }

      const data = JSON.parse(responseText);

      const totalProfit =
        data?.results?.[0]?.tables?.[0]?.rows?.[0]?.["[TotalProfit]"];

      console.log("TOTAL PROFIT VALUE:", totalProfit);
      console.log("FULL POWER BI RESULT:", JSON.stringify(data, null, 2));

      res.json({
        success: true,
        kpi: "Total Profit",
        value: totalProfit,
        readOnly: true,
        daxQuery,
      });
    } catch (error) {
      console.error("TOTAL PROFIT ERROR:", error);

      res.status(500).json({
        success: false,
        error: "Failed to retrieve Total Profit.",
        details: error.message,
      });
    }
  },
);

// =====================================================
// KPI - TOTAL ORDERS
// READ-ONLY QUERY
// =====================================================

app.get(
  "/api/workspaces/:workspaceId/datasets/:datasetId/kpi/total-orders",
  async (req, res) => {
    const { workspaceId, datasetId } = req.params;
    const connectionId = req.query.connectionId;

    console.log("=================================");
    console.log("KPI: TOTAL ORDERS");
    console.log("Workspace ID:", workspaceId);
    console.log("Dataset ID:", datasetId);
    console.log("Connection ID:", connectionId);
    console.log("=================================");

    if (!connectionId) {
      return res.status(400).json({
        success: false,
        error: "Connection ID is missing.",
      });
    }

    const connection = powerBiConnections.get(connectionId);

    if (!connection) {
      return res.status(401).json({
        success: false,
        error: "Invalid or expired PRISM AI connection.",
      });
    }

    if (!connection.authenticated || !connection.accessToken) {
      return res.status(401).json({
        success: false,
        error: "Power BI authentication required.",
      });
    }

    try {
      const daxQuery = `
        EVALUATE
        ROW(
          "TotalOrders",
          COUNTROWS('Orders')
        )
      `;

      console.log("Executing Total Orders DAX:");
      console.log(daxQuery);

      const response = await fetch(
        `${POWER_BI_API}/groups/${workspaceId}/datasets/${datasetId}/executeQueries`,
        {
          method: "POST",

          headers: {
            Authorization: `Bearer ${connection.accessToken}`,
            "Content-Type": "application/json",
          },

          body: JSON.stringify({
            queries: [
              {
                query: daxQuery,
              },
            ],

            serializerSettings: {
              includeNulls: true,
            },
          }),
        },
      );

      const responseText = await response.text();

      console.log("Total Orders API status:", response.status);
      console.log("Total Orders response:", responseText);

      if (!response.ok) {
        return res.status(response.status).json({
          success: false,
          error: "Unable to calculate Total Orders.",
          powerBiStatus: response.status,
          details: responseText,
        });
      }

      const data = JSON.parse(responseText);

      const totalOrders =
        data?.results?.[0]?.tables?.[0]?.rows?.[0]?.["[TotalOrders]"];

      console.log("TOTAL ORDERS VALUE:", totalOrders);
      console.log("FULL POWER BI RESULT:", JSON.stringify(data, null, 2));

      res.json({
        success: true,
        kpi: "Total Orders",
        value: totalOrders,
        readOnly: true,
        daxQuery,
      });
    } catch (error) {
      console.error("TOTAL ORDERS ERROR:", error);

      res.status(500).json({
        success: false,
        error: "Failed to retrieve Total Orders.",
        details: error.message,
      });
    }
  },
);

// =====================================================
// PRISM AI - LLM ASK
// =====================================================

app.post("/api/ask", async (req, res) => {
  try {
    const { question, context } = req.body;

    console.log("=================================");
    console.log("PRISM AI LLM REQUEST");
    console.log("Question:", question);
    console.log("=================================");

    // -----------------------------------------
    // VALIDATE QUESTION
    // -----------------------------------------

    if (!question || question.trim() === "") {
      return res.status(400).json({
        success: false,
        error: "Question is required.",
      });
    }

    // -----------------------------------------
    // SYSTEM PROMPT
    // -----------------------------------------

    const systemPrompt = `
You are Prism AI, an intelligent Business Intelligence assistant.

You help users understand and analyze data from their Power BI semantic model.

Your job is to answer the user's business question using ONLY the Power BI data and context provided to you.

Rules:

1. Give clear and concise answers.
2. Use the provided Power BI data whenever relevant.
3. Never invent numbers, dates, products, regions, or business facts.
4. If the provided data is insufficient to answer the question, clearly say that more Power BI data is required.
5. Explain calculations and comparisons in simple business language.
6. Preserve numerical values accurately.
7. If the user asks for a comparison, clearly identify both values and the difference.
8. If the user asks for a percentage change, calculate it accurately when the required values are available.
9. If the user asks for trends, summarize the trend using the provided data.
10. If the user asks a question unrelated to the available Power BI data, explain that the information is not available in the current semantic model.
11. Do not claim that you queried Power BI unless the context explicitly contains the query result.
12. Do not make assumptions about missing data.

You are a read-only BI assistant.
`;

    // -----------------------------------------
    // USER PROMPT
    // -----------------------------------------

    const userPrompt = `
User Question:
${question}

Power BI Context:
${context || "No Power BI context was provided."}
`;

    // -----------------------------------------
    // CALL QWEN3
    // -----------------------------------------

    const response = await llmClient.chat.completions.create({
      model: LLM_MODEL,

      messages: [
        {
          role: "system",
          content: systemPrompt,
        },
        {
          role: "user",
          content: userPrompt,
        },
      ],

      max_tokens: 300,
    });

    const answer = response?.choices?.[0]?.message?.content;

    if (!answer) {
      throw new Error("LLM returned an empty response.");
    }

    console.log("PRISM AI ANSWER:");
    console.log(answer);
    console.log("=================================");

    // -----------------------------------------
    // RETURN ANSWER
    // -----------------------------------------

    res.json({
      success: true,
      answer,
      model: LLM_MODEL,
    });
  } catch (error) {
    console.error("=================================");
    console.error("PRISM AI LLM ERROR");
    console.error(error);
    console.error("=================================");

    res.status(500).json({
      success: false,
      error: "LLM request failed.",
      details: error.message,
    });
  }
});

app.listen(PORT, () => {
  console.log("=================================");

  console.log("BI ASSISTANT BACKEND RUNNING");

  console.log(`http://localhost:${PORT}`);

  console.log("=================================");
});
