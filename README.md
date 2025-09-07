# Gmail DocLog One

## Overview

Gmail DocLog One is a Google Workspace add-on for Gmail that helps you automatically log information about email attachments and links into a Google Sheet. By simply applying a specific Gmail label to your emails, you can trigger the add-on to extract key data and organize it for you, creating a clean, searchable log in your Google Drive.

This tool is perfect for anyone who needs to track documents, links, or important attachments sent or received via email, such as project managers, administrators, or anyone handling digital receipts and invoices.

## How it Works

The add-on operates from the Gmail sidebar and follows a simple, user-friendly workflow:

1.  **Initial Setup**: The first time you run the add-on, it automatically searches for a Google Sheet named **"Gmail DocLog Sheet"** in your Google Drive.
    *   If it finds exactly one, it will use it.
    *   If it doesn't find one, it will create a new sheet with that name for you. The sheet will contain a "Data" tab with the necessary headers.
    *   If it finds multiple sheets with the same name, it will ask you to manually provide the correct Sheet ID or URL to resolve the ambiguity.

2.  **Label Configuration**: The add-on uses a Gmail label to identify which emails to process.
    *   By default, it looks for the label `---`.
    *   If this label doesn't exist, the add-on will prompt you to create it.
    *   You can change the label to any existing or new label you prefer directly from the add-on's settings.

3.  **Processing Emails**:
    *   Apply your chosen label to any email thread in Gmail that you want to log.
    *   Open the add-on in the sidebar and click the **"Process Now"** button.
    *   The add-on will then scan all messages in the labeled threads. For each message, it extracts all attachments and any links found in the email body.
    *   For each attachment or link, it adds a new row to the Google Sheet. A separate row is created for each recipient of the email, linking them to the specific item.
    *   Once a thread has been processed, the add-on removes the label to prevent it from being processed again.

## Features

*   **Automated Data Extraction**: Automatically pulls attachment names and URLs from emails.
*   **Google Sheet Integration**: Logs all extracted data neatly into a Google Sheet.
*   **Automatic Setup**: Creates the required Google Sheet if it doesn't already exist.
*   **Customizable Label**: Use the default `---` label or configure any other Gmail label.
*   **Simple User Interface**: All actions are controlled from a simple card interface in the Gmail sidebar.
*   **Prevents Duplicates**: Removes the label after processing to ensure emails are only logged once.
*   **Manual Override**: Provides options to reset the sheet configuration or manually select a sheet if needed.

## Setup and Usage

1.  **Install the Add-on**: Install Gmail DocLog One from the Google Workspace Marketplace.
2.  **Open the Add-on**: Open Gmail and click the add-on icon in the right-hand sidebar.
3.  **Authorize Permissions**: The first time you open it, you will be prompted to authorize the necessary permissions. See the "Permissions Explained" section for details.
4.  **Initial Configuration**:
    *   The add-on will guide you through setting up the Google Sheet. In most cases, this is fully automatic.
    *   It will also help you set up the Gmail label (default: `---`).
5.  **Start Logging**:
    *   Find an email with attachments or links you want to log.
    *   Apply the configured label (e.g., `---`) to the email thread.
    *   Click the **"Process Now"** button in the add-on sidebar.
    *   Open the "Gmail DocLog Sheet" in your Google Drive to see the logged data.

## Permissions Explained

This add-on requires a few permissions to function correctly. Here’s why each one is needed:

*   `https://www.googleapis.com/auth/gmail.addons.execute`: Base permission to run as a Gmail add-on.
*   `https://www.googleapis.com/auth/gmail.readonly`: To read the content of emails you label, including headers, attachments, and links.
*   `https://www.googleapis.com/auth/gmail.modify`: To remove the label from emails after they have been processed.
*   `https://www.googleapis.com/auth/gmail.send`: Used for the feature that allows sending a draft email after it has been processed.
*   `https://www.googleapis.com/auth/userinfo.email`: To identify you as the user.
*   `https://www.googleapis.com/auth/spreadsheets`: To create and write data to the "Gmail DocLog Sheet".
*   `https://www.googleapis.com/auth/drive`: To find or create the "Gmail DocLog Sheet" in your Google Drive.

## Configuration

You can customize the following settings from the add-on's main card:

*   **Gmail Label**: In the "Settings" section, you can change the name of the Gmail label the add-on uses for processing.
*   **Target Sheet**: You can reset the connection to the Google Sheet by clicking the "Reset" button. The add-on will then re-run the search/create process on its next launch.
