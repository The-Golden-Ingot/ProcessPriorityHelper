# Process Priority Helper (WPF)

A simple WPF tool to manage persistent CPU, I/O, and memory page priority overrides for processes on Windows 10/11.

It writes to:
```
HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\<program.exe>\PerfOptions
```
## Features
- View all current PerfOptions overrides (CPU, I/O, page priority)
- Add/edit overrides with validation and optional advanced values
- Remove overrides (values only or delete the whole PerfOptions key)
- Always uses the 64‑bit registry view on 64‑bit Windows
- Safe relaunch into elevated, 64‑bit, STA desktop host to ensure WPF reliability

## Requirements
- Windows 10/11
- PowerShell 5.1 (the script relaunches into 64‑bit powershell.exe with `-STA`  automatically)
- Administrator rights (writes to HKLM)

## Install / Run
- **Run the packaged executable**
  - Download [ProcessPriorityHelper.exe](https://github.com/The-Golden-Ingot/ProcessPriorityHelper/releases/download/1/ProcessPriorityHelper.exe)
  - Double-click it, and approve the UAC prompt.
  - If Microsoft SmartScreen blocks the executable,
    - Click "More info" in the SmartScreen warning dialog.
    - Click "Run anyway" to allow the executable to run.

- **Run the PowerShell script**
  - Download [ProcessPriorityHelper.ps1](https://github.com/The-Golden-Ingot/ProcessPriorityHelper/releases/download/1/ProcessPriorityHelper.ps1)
  - Recommended launch (temporary bypass, no policy changes persist):
    ```powershell
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\ProcessPriorityHelper.ps1"
    ```
    The script will relaunch itself as elevated, 64‑bit, and STA if needed.

### Optional logging
```powershell
.\ProcessPriorityHelper.ps1 -LogPath "C:\ProgramData\ProcessPriorityHelper\pph.log"    # Write a UTF‑8 text log
.\ProcessPriorityHelper.ps1 -Quiet                                                     # Suppress console logging
```

## Why relaunch to 64‑bit, STA, elevated?
- 64‑bit: guarantees access to the 64‑bit registry view
- STA: WPF requires STA for reliable UI
- Elevated: writes to HKLM require Administrator

## How it works
- Reads and writes DWORD values under `HKLM\...\IFEO\<exe>\PerfOptions` :
  - `CpuPriorityClass`  (1=Idle, 2=Normal, 3=High, 4=Realtime, 5=Below Normal, 6=Above Normal)
  - `IoPriority`  (0=Very Low, 1=Low, 2=Normal, 3=High)
  - `PagePriority`  (1=Low .. 5=High)
- Changes take effect when the process next starts

## Packaging (optional)
- Convert to EXE:
  ```powershell
  Install-Module ps2exe -Scope CurrentUser
  Invoke-PS2EXE .\ProcessPriorityHelper.ps1 .\ProcessPriorityHelper.exe -title "Process Priority Helper" -version 1.0.0
  ```

## License
[MIT](LICENSE)