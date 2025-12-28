# md2ADUC
Render a simulated Active Directory Users and Computers tree from a Markdown unordered list
<img width="808" height="716" alt="image" src="https://github.com/user-attachments/assets/ae610086-c24d-48ba-8d20-4103e7471c46" />


A vibe-coded ADUC (Active Directory Users and Computers) simulator that converts markdown lists into a realistic Windows-style tree view interface. Here's what the tool provides:

### Key Features:
- "Authentic" ADUC Interface: Mimics the classic Windows MMC console with title bar, toolbar, and tree view
- Markdown Parsing: Takes indented markdown lists with 2-space indentation
- Object Type Support: Recognizes these types in brackets:
  - [user] - Shows user icon 👤
  - [computer] - Shows computer icon 💻
  - [group] - Shows group icon 👥
  - [contact], [printer], [share], [policy], [container] - Various other icons
  - No bracket = OU (folder) 📁

### How to Use:
1. Open the HTML file in any modern browser
2. Enter your markdown in the left panel using this format:
```markdown
   - Root OU
     - Child OU
       - User Name [user]
       - Computer Name [computer]
```
3. Click "Render Tree" to generate the ADUC view
4. Click "Export as Image" to save as PNG for PowerPoint

The interface includes:
- Expandable/collapsible tree nodes
- Proper Windows-style window chrome
- Toolbar buttons (visual only)
- Status bar showing object count
- High-resolution export (2x scale for clarity)
- Save and Open

The tool loads with a sample structure showing a typical domain layout with Domain Controllers, Corporate OUs, and Branch Offices. You can modify the markdown to match your exact AD structure for documentation or presentations.
The exported PNG will have a clean white background perfect for inserting into PowerPoint slides.

The file operations work entirely in the browser using the HTML5 File API - no server required. Your markdown files can be version controlled, shared with colleagues, or built up as a library of standard structures for documentation.

This makes it perfect for:
- Building documentation templates
- Creating training materials
- Designing AD structures before implementation
- Maintaining a library of common OU structures
- Creating weird Active Directory themed PowerPoint slides

If you want to document an existing Active Directory environment, the complementary [PowerShell scripts in the AD2Markdown directory](./AD2Markdown/README.md) will be helpful.