\# Vibe coded ADS tool for quickly looking at PLC variables with minimal dependency.



\## Disclaimer



This guide is a personal project and not a peer-reviewed publication or sponsored document. It is provided “as is,” without any warranties—express or implied—including, but not limited to, accuracy, completeness, reliability, or suitability for any purpose. The author(s) shall not be held liable for any errors, omissions, delays, or damages arising from the use or display of this information.



All opinions expressed are solely those of the author(s) and do not necessarily represent those of any organization, employer, or other entity. Any assumptions or conclusions presented are subject to revision or rethinking at any time.



Use of this information, code, or scripts provided is at your own risk. Readers are encouraged to independently verify facts. This content does not constitute professional advice, and no client or advisory relationship is formed through its use.



\## Description



A single PowerShell script for searching, reading, writing and charting TwinCAT PLC variables on a PC that only has the TwinCAT ADS runtime installed. No XAE, Python or extra DLLs: it calls `TcAdsDll.dll` directly.

**Features**

- Pick a target from the ADS routes already on the PC (TC3 `StaticRoutes.xml` / TC2 registry), or type a Net ID.
- Search the uploaded symbols, including struct / FB members and small arrays.
- Read and write by name, with IEC value formats (`TRUE`, `16#FF`, `T#1s500ms`, `DT#...`, enum names, `hex:` for raw bytes).
- Watch list with live update, plus a 10 s live chart of the selected variable.
- Dark and light themes (follows the Windows setting by default).

**Usage**

Copy `AdsTool.ps1`, `AdsTool.bat` and `AdsTool-Console.bat` into one folder on the target PC.

- `AdsTool.bat`: GUI.
- `AdsTool-Console.bat`: console mode (`find`, `read`, `write`, `watch`, `routes`, ...).

Optional arguments (passed through by both .bat files): `-NetId <Net ID or route name>`, `-Port 851` (TC2: 801), `-Theme Auto|Dark|Light`, `-AdsDll <path to TcAdsDll.dll>`, `-TimeoutMs`, `-MaxDepth`, `-MaxArray`.

Remote targets need an ADS route on the PC first.

