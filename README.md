# Installing the Word Template (.dotm)

This README provides instructions for installing the Word template (.dotm) file to your Microsoft Office Custom Templates folder.

## Prerequisites

- Microsoft Word (any recent version)
- The .dotm template file from this repository

## Installation Instructions

### Windows

1. Download the .dotm file from this repository
2. Locate your Microsoft Word Custom Templates folder. The default location is:
   `C:\Users\YourUsername\AppData\Roaming\Microsoft\Templates`
3. Copy the .dotm file to this folder
4. Restart Microsoft Word if it's already running

### MacOS

1. Download the .dotm file from this repository
2. Locate your Microsoft Word Custom Templates folder. The default location is:
   `~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates`
3. If the Templates folder doesn't exist, you may need to create it
4. Copy the .dotm file to this folder
5. Restart Microsoft Word if it's already running

### Alternative Installation Method

1. Open Microsoft Word
2. Click on **File** > **Options** (Windows) or **Word** > **Preferences** (Mac)
3. On Windows: Select **Advanced** from the left sidebar, scroll down and click on **File Locations**
   On Mac: Click on **File Locations**
4. Look for "User templates" and click **Modify** (Windows) or **Change** (Mac)
5. This will open the folder where you should place the .dotm file
6. Copy the .dotm file to this folder
7. Restart Microsoft Word

## Using the Template

1. Open Microsoft Word
2. Click on **File** > **New** (Windows) or **File** > **New from Template** (Mac)
3. Click on **Personal** or **Custom** tab (depending on your Word version)
4. Select the template from the available options

## Importing the Template to an Existing Document

### Method 1: Using the Organizer

1. Open the existing Word document where you want to import the template
2. Press **Alt+F11** (Windows) or **Option+F11** (Mac) to open the Visual Basic Editor
3. In the VBE, click on **Tools** > **Organizer**
4. In the Organizer dialog, click on the **Styles** tab
5. On the right side, click **Close File** and then **Open File**
6. Browse to and select the .dotm template file
7. Select the styles you want to copy from the template (right side) to your document (left side)
8. Click **Copy** to import the selected styles
9. Close the Organizer and the VBE

### Method 2: Attaching the Template

1. Open the existing Word document
2. Click on **File** > **Options** (Windows) or **Word** > **Preferences** (Mac)
3. Select **Add-ins** (Windows) or **Templates and Add-ins** (Mac)
4. Click on **Templates** tab (Windows) or ensure you're on the Templates tab (Mac)
5. Click on **Attach** button
6. Browse to and select the .dotm template file
7. Check the box for "Automatically update document styles" if you want styles to update
8. Click **OK**

### Method 3: Copy and Paste Styles

1. Create a new document based on the .dotm template
2. Press **Ctrl+Alt+Shift+S** (Windows) or **Command+Option+Shift+S** (Mac) to open the Styles pane
3. For each style you want to import, right-click on it and select **Select All X Instance(s)** where X is the number of instances
4. Copy the selected content (**Ctrl+C** / **Command+C**)
5. Go to your existing document and paste (**Ctrl+V** / **Command+V**)
6. The styles will be imported along with the content
7. You can then modify or delete the pasted content while keeping the imported styles

## Troubleshooting

If you don't see the template after installation:
- Make sure you've placed the .dotm file in the correct folder
- Check if you have sufficient permissions to access the templates folder
- Restart Microsoft Word completely
- On Mac, you may need to check if the Library folder is hidden (press Cmd+Shift+. to show hidden files)

## Support

If you encounter any issues with the template installation or usage, please open an issue in this repository.