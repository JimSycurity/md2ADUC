# AD Object Class Mapping Reference







## Overview



This document explains how Active Directory object classes are mapped to markdown markers for the ADUC Visualization tool.







## Object Class Mappings







### AD Object Classes → Markdown Markers







| AD ObjectClass | Normalized Type | Markdown Output | Icon in Visualizer |



|----------------|-----------------|-----------------|-------------------|



| user | user | `Name [user]` | 👤 |



| computer | computer | `Name [computer]` | 💻 |



| group | group | `Name [group]` | 👥 |



| organizationalUnit | organizationalUnit | `Name` | 📁 |



| container | container | `Name` | 📁 |



| domainDNS | domainDNS | `Name` | 📁 |



| contact | contact | `Name [contact]` | 📇 |



| printQueue | printer | `Name [printer]` | 🖨️ |



| volume | share | `Name [share]` | 📤 |



| groupPolicyContainer | policy | `Name [policy]` | 📋 |



| msDS-ManagedServiceAccount | computer | `Name [computer]` | 💻 |



| msDS-GroupManagedServiceAccount | computer | `Name [computer]` | 💻 |







### Special Cases







1. **Managed Service Accounts**: sMSA, gMSA, and other `*ServiceAccount*` classes are mapped to `[computer]` because they are computer-derived identities



2. **Computer Names**: Trailing `$` is automatically removed (e.g., `COMPUTER$` → `COMPUTER`)



3. **Domain Root**: Domain components (DC=domain,DC=com) are combined into a single domain name (domain.com)



4. **No Markers**: OUs, containers, and domains don't get type markers - they display as folders







## Common Issues and Solutions







### Issue: Objects not displaying with correct icons



**Cause**: Raw AD objectClass values being passed through without normalization



**Solution**: The Convert-ADCsvToMarkdown.ps1 script now properly normalizes all object classes







### Issue: Computer names showing with $



**Cause**: AD stores computer accounts with trailing $



**Solution**: The converter automatically strips the $ from computer names







### Issue: Service accounts not showing as computers



**Cause**: MSA and gMSA object classes not mapped



**Solution**: Managed service account object classes now map to [computer]







## Testing Object Class Mapping







Run the diagnostic script to verify mappings:



```powershell



.\Test-ObjectClassMapping.ps1



```







This will:



- Test all common AD object class mappings



- Analyze any CSV files in the directory



- Show how objects will appear in markdown



- Identify any unsupported object types







## Supported Markers in ADUC Visualizer







The HTML visualizer (`aduc-simulator.html`) recognizes these exact markers:



- `[user]` - User icon 👤



- `[computer]` - Computer icon 💻



- `[group]` - Group icon 👥



- `[contact]` - Contact icon 📇



- `[printer]` - Printer icon 🖨️



- `[share]` - Share icon 📤



- `[policy]` - Policy icon 📋



- `[container]` - Container icon 📦



- (no marker) - Folder icon 📁/📂







## Workflow for Proper Conversion







1. **Export from AD with proper object classes:**



   ```powershell



   .\Export-ADToCSV.ps1 -OutputFile "ADObjects.csv"



   ```







2. **Verify object class mapping:**



   ```powershell



   .\Test-ObjectClassMapping.ps1



   ```







3. **Convert to markdown:**



   ```powershell



   .\Convert-ADCsvToMarkdown.ps1 -CsvFile "ADObjects.csv" -OutputFile "structure.md"



   ```







4. **Load in visualizer:**



   Open `aduc-simulator.html` and load the markdown file







## Manual Markdown Format







If creating markdown manually, use this format:



```markdown



- domain.com



  - Domain Controllers



    - DC01 [computer]



    - DC02 [computer]



  - Users (OU - no marker)



    - John Doe [user]



    - Jane Smith [user]



  - Computers (OU - no marker)



    - DESKTOP-01 [computer]



    - LAPTOP-02 [computer]



  - Groups (OU - no marker)



    - Domain Admins [group]



    - Finance Team [group]



  - Service Accounts



    - svc_backup [user]



    - svc_web [user]



```







## Validation







After conversion, check the output:



1. Users should have `[user]` marker



2. Computers should have `[computer]` marker (no $)



3. Groups should have `[group]` marker



4. OUs should have NO marker



5. Domain root should have NO marker







If any objects show with their raw objectClass (like `[domainDNS]` or `[msDS-ManagedServiceAccount]`), the mapping has failed and needs to be debugged.