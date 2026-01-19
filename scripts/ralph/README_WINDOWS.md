# Ralph on Windows - Mailtool MCP SDK Migration

## Why Windows?

Mailtool **must run on Windows** because:
1. **pywin32 dependency** - Only installs on Windows (platform-specific in pyproject.toml)
2. **COM automation** - Outlook bridge requires Windows COM objects via pywin32
3. **MCP server** - Must connect to real Outlook instance on Windows

All 50 user stories require Windows execution for testing and validation.

## Setup

### 1. Copy Ralph to Windows

From WSL2:
```bash
cd /home/sam/dev/mailtool
cp -r scripts/ralph /mnt/c/dev/mailtool/scripts/
```

Or from Windows PowerShell:
```powershell
Copy-Item -Recurse \\wsl.localhost\Ubuntu\home\sam\dev\mailtool\scripts\ralph C:\dev\mailtool\scripts\
```

### 2. Install Prerequisites

**On Windows (PowerShell):**
```powershell
# Install Claude Code CLI (if not already installed)
npm install -g @anthropic-ai/claude-code

# Verify installation
claude --version

# Install uv (Python package manager)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Verify installation
uv --version

# Install jq for JSON parsing (optional, PowerShell has ConvertFrom-Json)
winget install jqlang.jq
```

### 3. Set Environment Variables

```powershell
# Add to System Environment Variables or set per-session
$env:ANTHROPIC_API_KEY = "your-api-key-here"

# Or add to Windows System Environment Variables for persistence
# Settings → Edit environment variables for your account → New
# Variable name: ANTHROPIC_API_KEY
# Variable value: your-api-key-here
```

## Usage

### Start Ralph (PowerShell)

```powershell
cd C:\dev\mailtool\scripts\ralph

# Run with default 10 iterations
.\ralph.ps1

# Or specify custom iteration count
.\ralph.ps1 -MaxIterations 20

# Or use the batch wrapper
ralph.bat
ralph.bat 20
```

### Stop Gracefully

Press `s` at any time to stop after current iteration completes. You'll be prompted to confirm.

## Windows-Specific Considerations

### Execution Policy

If you get "execution of scripts is disabled" error:
```powershell
# Temporarily bypass (current session)
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Or allow permanently (requires admin)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
```

### Background Jobs

Ralph uses PowerShell background jobs to listen for stop signals. These are automatically cleaned up on exit.

### Path Handling

- Accepts both Windows paths (`C:\dev\mailtool`) and WSL paths (`\\wsl.localhost\Ubuntu\home\...`)
- JSON parsing uses `ConvertFrom-Json` (built into PowerShell)

### COM Object Handling

The MCP server tests will:
1. Start Outlook on Windows (must be running)
2. Create COM objects via pywin32
3. Test all 23 tools with real Outlook data

## Troubleshooting

### "claude command not found"
```powershell
# Add npm to PATH or use full path
$env:Path += ";C:\Users\YourUsername\AppData\Roaming\npm"
```

### "uv command not found"
```powershell
# uv should be in PATH after installation
# If not, add: C:\Users\YourUsername\.local\bin
$env:Path += ";C:\Users\YourUsername\.local\bin"
```

### PowerShell version
```powershell
# Check version
$PSVersionTable.PSVersion

# PowerShell 7+ recommended for best compatibility
# Install from: https://github.com/PowerShell/PowerShell/releases
```

### Outlook not accessible
1. Start Outlook on Windows before running tests
2. Ensure Outlook profile is configured
3. Check COM permissions in Outlook Trust Center

## Development Workflow

### Recommended Windows Setup

```powershell
# 1. Clone to Windows path (not WSL)
cd C:\dev
git clone <repo> mailtool

# 2. Copy Ralph scripts
cd C:\dev\mailtool\scripts
# Copy ralph directory here

# 3. Run Ralph
cd C:\dev\mailtool\scripts\ralph
.\ralph.ps1
```

### Why Not Run from WSL?

**Ralph CANNOT run from WSL for this project because:**
1. `pywin32` won't install on Linux (platform-specific dependency)
2. COM objects don't exist on Linux
3. All tests require Windows Outlook
4. MCP server must run on Windows

The WSL2 wrappers (`outlook.sh`, `run_tests.sh`) exist for manual testing, but Ralph's autonomous execution needs direct Windows access.

## File Structure

```
C:\dev\mailtool\scripts\ralph\
├── ralph.ps1           # Main PowerShell script
├── ralph.bat           # Batch wrapper
├── prompt.md           # Ralph's agent instructions
├── prd.json            # Product requirements (50 user stories)
├── progress.txt        # Progress log
├── .last-branch        # Current branch tracking
└── archive/            # Previous runs (auto-archived)
```

## Progress Tracking

Ralph automatically:
1. Archives previous runs when branch changes
2. Updates `prd.json` to mark completed stories
3. Appends detailed progress to `progress.txt`
4. Commits changes after each completed story
5. Consolidates patterns in progress.txt header

View progress anytime:
```powershell
Get-Content C:\dev\mailtool\scripts\ralph\progress.txt | Select-Object -Last 50
```

## Next Steps

1. Ensure Outlook is running on Windows
2. Open PowerShell in `C:\dev\mailtool\scripts\ralph`
3. Run `.\ralph.ps1`
4. Ralph will autonomously implement all 50 user stories
5. Check `progress.txt` for detailed status

Estimated time: 3-4 weeks for full migration (50 stories × ~4 hours/story including testing)
