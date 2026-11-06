#  Power BI BI Assistant Backend

An AI-powered backend service for PRISM AI, a conversational Power BI assistant. This Node.js application connects Microsoft Power BI semantic models with Google Gemini AI, allowing users to ask natural-language questions and receive data-driven answers, DAX queries, KPIs, and visualization instructions.

## 1. Project Overview

PRISM AI is designed to simplify Power BI data exploration through a conversational interface.

The backend acts as the communication layer between the Power BI custom visual, Microsoft Entra ID, Microsoft Power BI REST API, and Google Gemini API.

### Key Features

- Microsoft account authentication using Microsoft Entra ID.
- Secure OAuth 2.0 authorization-code authentication flow.
- Power BI workspace retrieval.
- Semantic model and dataset selection.
- Automatic semantic model schema discovery.
- Natural-language-to-DAX generation using Google Gemini.
- Read-only DAX query execution against Power BI semantic models.
- AI-generated answers based on actual Power BI query results.
- KPI retrieval, including Total Sales, Total Profit, and Total Orders.
- Visualization planning for KPI, bar, line, pie, and table outputs.
- REST API communication with the Power BI custom visual.

## 2. System Architecture

PRISM AI consists of two separately maintained projects:

1. **Frontend:** Power BI Custom Visual.
2. **Backend:** Node.js application hosted on Render.

### Architecture Flow

```text
               USER
                 |
                 v
       POWER BI CUSTOM VISUAL
          (Frontend)
                 |
                 | HTTP API Requests
                 v
          RENDER PLATFORM
       (Node.js Backend Server)
                 |
          +------+------+
          |             |
          v             v
   MICROSOFT ENTRA    GOOGLE GEMINI
       ID API             API
          |             |
          v             v
    Power BI Token    AI Processing
          |             |
          v             |
    POWER BI REST API   |
          |             |
          v             |
    Semantic Model      |
       / DAX Query       |
          |             |
          +------+------+
                 |
                 v
          BACKEND RESPONSE
                 |
                 v
       POWER BI CUSTOM VISUAL
                 |
                 v
        ANSWERS / KPIs /
        CHARTS / TABLES
```



## System Architecture

The following diagram illustrates how the BI Assistence frontend, backend, Microsoft Entra ID, Google Gemini API, and Microsoft Power BI work together.

![BI Assistence Architecture](assets/Bot%20architecture.png)

### How Frontend and Backend Communicate

The Power BI custom visual sends HTTP requests to the backend using its configured backend URL.

Current production backend:

```text
https://bi-assistance-chat-bot.onrender.com
```

The backend is hosted on Render, which provides a public HTTPS endpoint.

The frontend uses this URL to communicate with backend API endpoints.

For example:

```text
Frontend
   |
   | POST /api/ask
   |
   v
Render Backend
   |
   | Processes the question
   | Calls Gemini API
   | Executes DAX using Power BI REST API
   |
   v
Returns JSON Response
   |
   v
Frontend displays the result
```

**Important:** The Power BI custom visual is not deployed on Render. It is packaged as a `.pbiviz` file and imported into Power BI. Render hosts only the backend service.

## 3. Technologies and Packages

### Technologies

- Node.js
- JavaScript
- Express.js
- Microsoft Entra ID
- Microsoft Authentication Library (MSAL)
- Microsoft Power BI REST API
- Google Gemini API
- Render
- GitHub

### Node.js Dependencies

The project dependencies are defined in `package.json`.

| Package | Purpose |
|---|---|
| `express` | Creates the backend server and REST API endpoints |
| `@azure/msal-node` | Microsoft Entra authentication and token acquisition |
| `@google/generative-ai` | Integrates Google Gemini AI |
| `axios` | Sends HTTP requests to Power BI APIs |
| `cors` | Handles cross-origin requests |
| `dotenv` | Loads environment variables |
| `express-session` | Manages Express sessions |
| `openai` | Additional AI client dependency; not currently used by the active Gemini implementation |

The project also uses Node.js built-in modules such as `crypto` and the built-in `fetch` API.

## 4. Prerequisites

Install the following tools before running the backend.

| Tool | Purpose |
|---|---|
| Node.js 18 or later | JavaScript runtime |
| npm | Dependency installation |
| Git | Clone and manage the repository |
| Visual Studio Code | Recommended development environment |
| Microsoft Entra access | Application registration and authentication configuration |
| Google AI Studio access | Generate Gemini API key |
| Render account | Backend cloud deployment |
| Power BI access | Test workspace and semantic model connectivity |

### Download Links

- Node.js: https://nodejs.org/
- Git: https://git-scm.com/
- Visual Studio Code: https://code.visualstudio.com/
- Microsoft Entra Admin Center: https://entra.microsoft.com/
- Google AI Studio: https://aistudio.google.com/
- Render: https://render.com/

## 5. Clone the Repository

Open PowerShell or a terminal.

Execute:

```bash
git clone https://github.com/Sahithi-mamidipally/BI-Assistance-Chat-Bot.git
```

Navigate into the backend directory:

```bash
cd BI-Assistance-Chat-Bot
```

## 6. Install Dependencies

Install the dependencies defined in `package.json`:

```bash
npm ci
```

Alternatively:

```bash
npm install
```

No need to install Express, MSAL, Gemini, Axios, or other dependencies individually.

## 7. Environment Configuration

The backend requires environment variables to connect to Microsoft Entra ID and Google Gemini.

Create a `.env` file in the repository root, alongside `server.js`.

### Required Environment Variables

```env
CLIENT_ID=your_microsoft_entra_client_id

CLIENT_SECRET=your_microsoft_entra_client_secret

TENANT_ID=your_microsoft_entra_tenant_id

REDIRECT_URI=http://localhost:3000/auth/callback

SESSION_SECRET=your_long_random_session_secret

GEMINI_API_KEY=your_google_gemini_api_key
```

### Environment Variable Description

| Variable | Description |
|---|---|
| CLIENT_ID | Client/application ID from Microsoft Entra |
| CLIENT_SECRET | Client secret generated in Microsoft Entra |
| TENANT_ID | Microsoft Entra directory tenant ID |
| REDIRECT_URI | OAuth authentication callback URL |
| SESSION_SECRET | Secret used to sign Express session cookies |
| GEMINI_API_KEY | API key generated through Google AI Studio |

**Security:**

- Never upload `.env` to GitHub.
- Never expose Gemini API keys in frontend code.
- Never share Microsoft client secrets.
- Use a unique, strong session secret.
- Add `.env` to `.gitignore`.

## 8. Microsoft Entra ID Configuration

Microsoft Entra ID is used to authenticate Power BI users and obtain access tokens for Power BI REST API requests.

### Application Registration

1. Open https://entra.microsoft.com/
2. Navigate to Identity → Applications → App registrations.
3. Select New registration.
4. Enter an application name, such as `PRISM AI BI Assistant`.
5. Configure the supported account types according to your organization.
6. Register the application.

### Collect Application Credentials

From the registered application, collect:

- Application (Client) ID.
- Directory (Tenant) ID.

Navigate to Certificates & secrets and create a client secret.

Store the generated secret value securely in your backend environment variables.

### Configure Authentication

Configure the application as a Web application with the following redirect URI:

Local development:

```text
http://localhost:3000/auth/callback
```

Render deployment:

```text
https://your-service-name.onrender.com/auth/callback
```

The redirect URI must exactly match the URI registered in Microsoft Entra.

### Configure Power BI API Permissions

The current backend requests these delegated permissions:

```text
Workspace.Read.All
Report.Read.All
Dataset.Read.All
```

These permissions allow the application to request access to the signed-in user's Power BI workspaces, reports, and datasets.

Add the required Power BI Service delegated API permissions in the Entra application registration.

Grant admin consent if required by your organization's consent policy.

**Note:** Users must also have the appropriate permissions to access the relevant Power BI workspaces and semantic models.

### Authentication Flow

```text
Power BI Custom Visual
          |
          v
Create Connection
          |
          v
Backend Generates Connection ID
          |
          v
Microsoft Login
          |
          v
Microsoft Entra Authentication
          |
          v
User Grants Required Consent
          |
          v
/auth/callback
          |
          v
Backend Exchanges Authorization Code
for Power BI Access Token
          |
          v
Connection Authenticated
          |
          v
Power BI Workspace Retrieval
```

The backend uses `@azure/msal-node` and the Microsoft authorization-code flow.

## 9. Google Gemini API Configuration

Google Gemini is currently used as the AI provider in PRISM AI.

The Gemini integration is implemented using:

```javascript
const { GoogleGenerativeAI } = require("@google/generative-ai");

const gemini = new GoogleGenerativeAI(
    process.env.GEMINI_API_KEY
);
```

Current configured model:

```javascript
const LLM_MODEL = "gemini-3.5-flash-lite";
```

The model is used for:

1. Understanding user questions.
2. Generating DAX queries from the semantic model schema.
3. Interpreting Power BI query results.
4. Generating concise natural-language answers.
5. Preparing visualization configuration.

### Generate Gemini API Key

1. Visit https://aistudio.google.com/
2. Sign in with your Google account.
3. Open the API keys section.
4. Create an API key.
5. Copy the key.
6. Add it to your local `.env` file or Render environment settings.

```env
GEMINI_API_KEY=your_google_gemini_api_key
```

The API key must remain on the backend.

## 10. Run the Backend Locally

After installing dependencies and configuring `.env`, start the application.

```bash
npm start
```

This executes:

```bash
node server.js
```

The backend uses:

```javascript
const PORT = process.env.PORT || 3000;
```

For local development, the default port is 3000.

### Expected Terminal Output

```text
=================================
MSAL CONFIGURATION
Client ID: [configured client ID]
Tenant ID: [configured tenant ID]
Redirect URI: http://localhost:3000/auth/callback
=================================

=================================
BI ASSISTANT BACKEND RUNNING
http://localhost:3000
=================================
```

The application should remain running without startup errors.

### Verify Backend Availability

Open:

```text
http://localhost:3000/
```

Expected response:

```text
BI Assistant backend is running successfully.
```

Test endpoint:

```text
http://localhost:3000/test
```

Expected JSON:

```json
{
  "success": true,
  "message": "TEST ROUTE WORKS"
}
```

Health endpoint:

```text
http://localhost:3000/api/health
```

Expected JSON structure:

```json
{
  "success": true,
  "message": "Backend reachable from Power BI visual",
  "sessionId": "generated-session-id",
  "authenticated": false
}
```

The session ID is generated dynamically. `authenticated: false` is normal before authentication.

## 11. Deploy the Backend on Render

Render hosts the backend and provides a public HTTPS URL.

### Step 1: Push Backend Code to GitHub

Make sure the following files are present in the repository:

```text
BI-Assistance-Chat-Bot/
│
├── server.js
├── package.json
├── package-lock.json
├── .gitignore
└── README.md
```

Do not upload `.env`.

### Step 2: Create a Render Web Service

1. Sign in to https://render.com/
2. Select New → Web Service.
3. Connect your GitHub account.
4. Select the backend repository.
5. Configure the service.

### Step 3: Configure Deployment Settings

Use the following settings:

| Setting | Value |
|---|---|
| Name | bi-assistance-chat-bot |
| Runtime | Node |
| Branch | master |
| Root Directory | Leave blank if server.js is at repository root |
| Build Command | `npm ci` |
| Start Command | `npm start` |

### Step 4: Configure Environment Variables

Navigate to the Render service's Environment section.

Add the following variables:

```text
CLIENT_ID
CLIENT_SECRET
TENANT_ID
REDIRECT_URI
SESSION_SECRET
GEMINI_API_KEY
```

For `REDIRECT_URI`, use the actual Render service URL:

```text
https://your-service-name.onrender.com/auth/callback
```

Make sure the same URL is registered in Microsoft Entra.

Render automatically provides the `PORT` environment variable.

### Step 5: Deploy

Select Deploy Web Service.

Render will:

1. Clone the GitHub repository.
2. Install Node.js dependencies.
3. Execute the configured start command.
4. Start the Express backend.
5. Provide a public HTTPS endpoint.

## 12. Verify Successful Render Deployment

After deployment, open the Render Logs section.

### Expected Startup Messages

Your current `server.js` should display:

```text
=================================
MSAL CONFIGURATION
=================================

=================================
BI ASSISTANT BACKEND RUNNING
http://localhost:<PORT>
=================================
```

**Note:** The current source prints `localhost` in its startup message even on Render. This is only a console message. Use the actual `onrender.com` URL shown in your Render dashboard to access the deployed service.

### Verify the Live Backend

Open your Render URL:

```text
https://your-service-name.onrender.com/
```

Expected response:

```text
BI Assistant backend is running successfully.
```

Test:

```text
https://your-service-name.onrender.com/test
```

Expected response:

```json
{
  "success": true,
  "message": "TEST ROUTE WORKS"
}
```

Health check:

```text
https://your-service-name.onrender.com/api/health
```

Expected:

```json
{
  "success": true,
  "message": "Backend reachable from Power BI visual",
  "sessionId": "generated-session-id",
  "authenticated": false
}
```

### Test Gemini Connectivity

Open:

```text
https://your-service-name.onrender.com/api/test-gemini
```

Expected response:

```json
{
  "success": true,
  "message": "Gemini connection successful"
}
```

The Gemini response may contain the exact requested text, with no additional content.

### Successful Authentication

When a user connects their Power BI account, Render logs should display messages similar to:

```text
NEW PRISM AI CONNECTION CREATED

STARTING MICROSOFT LOGIN

AUTH CALLBACK RECEIVED

USER AUTHENTICATED SUCCESSFULLY

PRISM AI CONNECTION UPDATED
```

The authentication callback should display:

```text
Power BI Connected Successfully!
```

These messages confirm that the corresponding operations have completed.

A successful backend health check alone does not prove that Microsoft authentication, Power BI permissions, and Gemini connectivity are all working. Test each separately.

## 13. Important Backend API Endpoints

The backend exposes several API endpoints used by the custom visual.

| Endpoint | Method | Purpose |
|---|---|---|
| `/` | GET | Backend availability |
| `/test` | GET | Basic backend test |
| `/api/health` | GET | Backend health and session status |
| `/api/test-gemini` | GET | Gemini connectivity test |
| `/api/create-connection` | GET | Generate a connection ID |
| `/login` | GET | Start Microsoft authentication |
| `/auth/callback` | GET | Receive OAuth authorization response |
| `/api/auth-status` | GET | Check authentication status |
| `/api/workspaces` | GET | Retrieve Power BI workspaces |
| `/api/workspaces/:workspaceId/datasets` | GET | Retrieve semantic models |
| `/api/workspaces/:workspaceId/reports` | GET | Retrieve Power BI reports |
| `/api/ask` | POST | Process natural-language questions |
| `/api/workspaces/:workspaceId/datasets/:datasetId/schema` | GET | Discover semantic model schema |
| `/api/workspaces/:workspaceId/datasets/:datasetId/metadata` | GET | Retrieve dataset metadata |

Additional endpoints support report pages, DAX testing, and KPI calculations.

## 14. How the AI Question Processing Works

When a user enters a question in PRISM AI:

### Step 1: Receive User Question

The frontend sends a POST request to:

```text
/api/ask
```

Example request:

```json
{
  "question": "What is the total sales?",
  "connectionId": "generated-connection-id",
  "workspaceId": "workspace-id",
  "datasetId": "dataset-id",
  "responseMode": "auto",
  "chartType": "auto"
}
```

### Step 2: Discover Semantic Model

The backend retrieves the selected semantic model's metadata using the Power BI REST API.

It discovers tables, columns, data types, and hidden-column information.

### Step 3: Generate DAX Using Gemini

The backend sends the semantic model schema and user question to Gemini.

Gemini generates a DAX query based on the actual schema.

### Step 4: Execute DAX

The backend executes the generated DAX query using:

```text
Power BI REST API
/groups/{workspaceId}/datasets/{datasetId}/executeQueries
```

### Step 5: Generate Final Answer

The actual Power BI query results are sent to Gemini to generate a natural-language response.

### Step 6: Generate Visualization Instructions

Gemini prepares visualization configuration according to the requested response mode and chart type.

Supported visualization types:

- KPI
- Bar chart
- Line chart
- Pie chart
- Table
- None

### Step 7: Return JSON

The backend returns:

```json
{
  "success": true,
  "question": "What is the total sales?",
  "answer": "Generated answer based on Power BI results.",
  "dax": "Generated DAX query",
  "data": [],
  "visualization": {
    "visualType": "kpi",
    "title": "Total Sales",
    "categoryField": null,
    "valueField": "TotalSales"
  }
}
```

The actual values and visualization configuration depend on the selected semantic model and Gemini response.

## 15. Important Project Files

```text
BI-Assistance-Chat-Bot/
│
├── server.js
├── package.json
├── package-lock.json
├── .env
├── .gitignore
└── README.md
```

| File | Responsibility |
|---|---|
| `server.js` | Main backend application, API endpoints, authentication, Power BI integration and Gemini processing |
| `package.json` | Dependency definitions and npm scripts |
| `package-lock.json` | Exact dependency version information |
| `.env` | Local environment variables and credentials; never commit |
| `.gitignore` | Prevents sensitive and unnecessary files from being committed |
| `README.md` | Project documentation and deployment instructions |

### Important npm Commands

| Command | Purpose |
|---|---|
| `npm ci` | Install locked dependencies |
| `npm install` | Install dependencies |
| `npm start` | Start backend using server.js |
| `node server.js` | Run backend directly |

The current `package.json` does not define a build script. The backend is a Node.js application that runs directly.

## 16. Connecting Your Own Power BI Custom Visual

If another developer deploys the backend using their own Render account, they must update the frontend backend URL.

The frontend and backend are maintained separately.

### Backend

Deploy the backend repository to Render.

Example:

```text
https://your-service-name.onrender.com
```

### Frontend

Open the custom visual project:

```text
bIAssistant/src/visual.ts
```

Update the backend URL configuration to:

```typescript
const BACKEND_URL =
    "https://your-service-name.onrender.com";
```

Use the actual configuration variable name in your frontend source.

After updating:

1. Install frontend dependencies.
2. Run the Power BI visual development server.
3. Test the connection.
4. Package the custom visual.
5. Import the updated `.pbiviz` file into Power BI.

This connects the frontend to the newly deployed backend.

**Important:** Updating the backend URL requires rebuilding and redistributing the custom visual. Deploying the backend alone does not automatically update an already imported Power BI visual.

## 17. Updating and Deploying Backend Changes

When making changes to `server.js`:

```bash
git pull origin master

# Make the required changes

npm install

npm start
```

After local testing:

```bash
git status

git add .

git commit -m "Update PRISM AI backend"

git push origin master
```

If Render is connected to the same GitHub repository and branch with automatic deployment enabled, pushing changes can trigger a new deployment.

Render installs dependencies and restarts the backend based on its deployment configuration.

## 18. Security and Deployment Considerations

- Never commit `.env` files.
- Never expose Microsoft client secrets or Gemini API keys.
- Keep environment variables in Render's Environment settings.
- Use HTTPS redirect URIs for production.
- Avoid logging OAuth authorization codes or access tokens.
- Use a strong, non-default `SESSION_SECRET`.
- Configure appropriate frontend origins and CORS policies.
- Use a persistent session store and connection-state store before scaling to multiple instances or relying on authentication across restarts.
- Validate AI-generated DAX before execution.
- Restrict Power BI access according to the user's permissions.

**Current implementation note:** Connection information is held in an in-memory `Map`, and Express uses its default in-memory session store. These are development-oriented choices. Authentication state can be lost when the process restarts, and these stores are not shared between multiple instances. A persistent session and connection store should be implemented for production-scale deployment.

The OAuth callback also currently logs query parameters. Remove that logging or sanitize it before production deployment so authorization codes are not exposed in service logs.

## 19. Official Documentation

- Microsoft Entra Admin Center: https://entra.microsoft.com/
- Microsoft Identity Platform: https://learn.microsoft.com/en-us/entra/identity-platform/
- Microsoft Power BI REST API: https://learn.microsoft.com/en-us/rest/api/power-bi/
- MSAL Node: https://learn.microsoft.com/en-us/entra/msal/javascript/node/
- Google Gemini API: https://ai.google.dev/gemini-api/docs
- Express.js: https://expressjs.com/
- Render Web Services: https://render.com/docs/web-services
- Render Deployment Guide: https://render.com/docs/deploy-node-express-app

---

## Project Repositories

**Backend Repository:**

https://github.com/Sahithi-mamidipally/BI-Assistance-Chat-Bot

**Frontend Repository:**

https://github.com/Sahithi-mamidipally/PowerBI-Custom-Visual

**Current Backend Hosting:**

https://bi-assistance-chat-bot.onrender.com

**Project:** PRISM AI – Power BI BI Assistant

**AI Provider:** Google Gemini API

**Authentication:** Microsoft Entra ID

**Backend Hosting:** Render
