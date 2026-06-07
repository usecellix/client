# Running Cellix in Excel Desktop

## One-Time Setup

### 1. Trust HTTPS Certificate
Before running the add-in, you need to trust the self-signed certificate for localhost:

```powershell
npx office-addin-dev-certs install
```

This will:
- Generate a self-signed certificate for localhost
- Install it in your Windows certificate store
- Allow Excel to load the add-in over HTTPS

**Note:** You only need to do this once per machine.

## Running the Add-in

You need to run **TWO commands in separate terminals**:

### Terminal 1: Start the Dev Server
```powershell
cd frontend
npm run dev
```

This starts Vite on https://localhost:3000

### Terminal 2: Launch Excel with Add-in
```powershell
cd frontend
npm start
```

This will:
- ✅ Automatically launch Excel desktop application
- ✅ Sideload the add-in (visible in the Home tab ribbon)
- ✅ Open the taskpane with your Cellix AI Assistant

## What to Expect

1. Excel will open automatically
2. Look for **"Show Taskpane"** button in the **Home** tab ribbon under "Commands Group"
3. Click it to open the Cellix AI Assistant sidebar
4. The sidebar will load from https://localhost:3000

## Stopping the Add-in

To remove the add-in from Excel:
```powershell
npm stop
```

## Alternative Commands

- **Start in desktop Excel:** `npm run start:desktop`
- **Start in web Excel:** `npm run start:web`
- **Validate manifest:** `npm run validate`

## Troubleshooting

### Certificate Issues
If you see "The content is blocked because it isn't signed by a valid security certificate":

1. **Install/Reinstall the certificate:**
```powershell
npx office-addin-dev-certs install --force
```

2. **Close Excel completely:**
   - Close all Excel windows
   - Check Task Manager to ensure no Excel processes are running

3. **Clear Excel cache (if error persists):**
   - Close Excel
   - Delete the folder: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Restart Excel

4. **Verify certificate is trusted:**
```powershell
npx office-addin-dev-certs verify
```

**Note:** The Vite dev server is configured to use the Office Add-in certificate automatically. If you still see errors, make sure the certificate is installed and Excel cache is cleared.

### Excel doesn't launch
- Make sure Excel is installed on your machine
- Close any existing Excel instances
- Try running `npm start` again

### Add-in not loading
- Verify dev server is running on https://localhost:3000
- Check that manifest.xml exists in the frontend folder
- Clear Excel cache: Close Excel, delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

### Port already in use
If port 3000 is already in use, you'll need to:
1. Change the port in `vite.config.ts`
2. Update all URLs in `manifest.xml`

