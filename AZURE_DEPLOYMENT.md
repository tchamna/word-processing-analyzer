# Azure Deployment Guide for Docx Formatting Analyzer

## Prerequisites
- Azure account with active subscription
- Azure CLI installed (optional, for command-line deployment)

## Deployment Steps

### 1. Create Azure App Service
1. Go to [Azure Portal](https://portal.azure.com)
2. Click "Create a resource" > "Web App"
3. Configure:
   - **Name**: Choose a unique name (e.g., `docx-analyzer-app`)
   - **Runtime stack**: Python 3.9 or later
   - **Region**: Select your preferred region
   - **App Service Plan**: Create new or use existing (B1 or higher recommended)
4. Click "Review + create" > "Create"

### 2. Configure the App Service
1. Go to your newly created App Service
2. Under "Settings" > "Configuration"
3. Add application settings:
   - **SCM_DO_BUILD_DURING_DEPLOYMENT**: `true`
   - **WEBSITE_RUN_FROM_PACKAGE**: `1` (optional, for faster deployments)

### 3. Deploy the Code
Choose one of the following methods:

#### Option A: GitHub Integration (Recommended)
1. In your App Service > "Deployment Center"
2. Select "GitHub" as source
3. Connect your GitHub account and select the repository
4. Choose branch (main/master)
5. Save - this will trigger automatic deployment

#### Option B: Manual Upload via FTP
1. In App Service > "Deployment Center" > "FTPS credentials"
2. Note the FTP details
3. Use an FTP client to upload all project files to `/site/wwwroot/`

#### Option C: Azure CLI
```bash
az webapp up --name <your-app-name> --resource-group <resource-group> --runtime "PYTHON:3.9" --location <location>
```

### 4. Configure Startup Command
In App Service > "Configuration" > "General settings":
- **Startup Command**: `streamlit run app.py --server.port %PORT% --server.address 0.0.0.0 --server.headless true`

### 5. Access Your App
Once deployed, your app will be available at: `https://<your-app-name>.azurewebsites.net`

## Troubleshooting

### Common Issues
1. **Port binding**: Ensure the startup command uses `%PORT%` for dynamic port assignment
2. **File uploads**: Azure App Service has file size limits (default 100MB)
3. **Memory**: Use at least B1 App Service Plan for adequate performance
4. **Dependencies**: Ensure all packages in requirements.txt are compatible with Azure's Python runtime

### Logs
Check application logs in Azure Portal > App Service > "Log stream" for debugging.

## Cost Considerations
- App Service has a free tier for testing
- B1 plan (~$13/month) recommended for production use
- Monitor usage to avoid unexpected costs

## Security Notes
- Consider adding authentication if the app will handle sensitive documents
- Use HTTPS (enabled by default on Azure App Service)
- Regularly update dependencies for security patches