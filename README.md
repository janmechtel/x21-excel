# X21 - Excel Add-in

An Excel VSTO add-in that provides AI-powered formula fixing and chat
functionality using AI.

## Prerequisites

- **Visual Studio 2022 Community Edition** with Office/SharePoint
  development workload be sure to select 2022 not 2026 or anything else
  <https://aka.ms/vs/17/release/vs_community.exe>
- **.NET Framework 4.8**
- Node.js
- Deno
- git
- nuget
- **Microsoft Excel 2016 or later**
- optional: Docker (optionally required for Semgrep security checks in pre-commit hooks)

## Setup Instructions

### Developement

For development run the 3 components locally Web UI, Deno server, and VSTO add-in. You can run them in any order but typically you would:

#### Web-UI

1. `cd ..\x21\x21\web-ui`
2. `npm install`
3. `npm run dev`

#### Deno Server

1. `cd  ..\x21\x21\deno-server`
2. `deno task start`

#### VSTO Add-in

1.  open `..\x21\X21.sln` in Visual Studio
2.  press F5 to start the add-in in Debug mode

### Next Steps

1. Publish locally and test the installer
2. Publish and sign to distribute the add-in to users
3. Setup the pre-commit hooks for code quality checks
4. See the .env-example for optional environment variables like Anthropic, Supabase and posthog



## Publish

### Publishing Prerequisites

- Install the `X21_TemporaryKey.pfx` (no password) before publishing
- Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
- run `nuget restore X21.sln`
- Publish to local with `.\Publish.ps1 -Environment Dev`


### Advanced Publish to Production (promote from staging)
- Publish to local with `.\Publish.ps1 -Environment Dev`
- Test
- Promote to staging with `.\Promote.ps1 -Environment Staging`
- Publish to local with `.\Publish.ps1 -Environment Staging`
- Merge back to dev

For production, you need to have the certificate installed on the machine and the hash of the certificate needs to match the value in Publish.ps1.

- for upload rclone must be installed for file deployment to Cloudflare R2 (see rclone
  setup below)

### Publishing and Signing
Alternative you could generate your own temporary certificate for signing the add-in:

1. Open **Visual Studio**
2. Go to **Project** → **X21 Properties**
3. Click on the **Signing** tab
4. Check **"Sign the ClickOnce manifests"**
5. Click **"Create Test Certificate..."**
6. Leave the password fields empty and click **OK**
7. This will generate `X21_TemporaryKey.pfx` in your project
8. Click **"Select from File..."** and choose the certificate you just created

Afterwards look for the thumbprint of the certificate in the certificate manager (`certmgr.msc`) and update the `PublishCore.psm1` script with the thumbprint value.

Then deploy this certificate to your users or get a real EV signing certificate, for example Azure Artifact Signing.

Things to check for if installation fails: Install location, Unblocking files, installing certificate to the right store (Trusted Publishers and Trusted Root Certification Authorities) 

### Git Operations

After successful publishing (but before deployment), the script will:

1. **Check Changes**: Verify that only the `.csproj` file has been modified
   (version increment)
2. **Commit Changes**: Automatically commit the version changes with a
   descriptive message
3. **Create Tag**: Create a Git tag in the format `vX.X.X.X` (e.g.,
   `v1.2.3.4`)
4. **Push Options**: Offer to push both the commit and tag to the remote
   repository

**Note**: If changes are detected in files other than `.csproj`, the script
will warn you and ask for confirmation before proceeding.

### Pre-commit Hooks

Install pre-commit hooks to run code quality checks before commits:

```bash
# Install pre-commit (requires Python)
pip install pre-commit

# Install the git hooks
pre-commit install

# Install pre-push hooks (for build checks)
pre-commit install --hook-type pre-push
```

The hooks will automatically run on `git commit` and validate:

- Deno server format/lint
- Web UI ESLint
- File formatting and validation
- Python linting (flake8, isort)
- Markdown linting
- Spell checking
- Security checks (Semgrep via Docker)

**Note**: Build checks (Web UI and VSTO Add-in) and security checks
(Semgrep) run on `git push` (pre-push hook) to keep commits fast.

#### rclone Setup for Cloudflare R2

- Download and install rclone: <https://rclone.org/downloads>
- Run rclone config and follow prompts:
  -- Run rclone config and follow prompts:
  -- Choose "n" for new remote
  -- Name it: x21
  -- Select S3 as the storage type
  -- Choose Cloudflare R2 as provider
  -- Leave region empty
  -- Set access_key_id and secret_access_key with credentials
  <https://dash.cloudflare.com/6bd46c67bb736bb693308127d72cd421/r2/api-tokens>

  ![image](https://github.com/user-attachments/assets/097fe5fb-4702-43f1-9112-21f7bef904ca)

-- Set endpoint to:
<https://6bd46c67bb736bb693308127d72cd421.r2.cloudflarestorage.com>

- Your resulting configuration should look like:

```[x21]
type = s3
provider = Cloudflare
access_key_id = <your_key>
secret_access_key = <your_key>
endpoint = https://6bd46c67bb736bb693308127d72cd421.r2.cloudflarestorage.com
```
