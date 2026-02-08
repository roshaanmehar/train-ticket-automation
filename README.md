# Train Ticket Collector (Google Apps Script)

## Overview

I needed a reliable way to keep track of all of my train ticket receipts.

Originally, doing this manually was a painful process. Every time I bought a ticket, I had to:

- Open the train ticket app on my phone  
- Find the ticket / receipt  
- Download the PDF to my mobile  
- Send it to myself on WhatsApp  
- Move it from WhatsApp onto my PC  
- Store it locally in a folder

It worked, but it was messy, time-consuming, and honestly a pretty ridiculous workflow for something I do regularly.

So I decided to automate the whole process.

---

## Why I Built This

My first idea was to find an official API from the ticket provider so I could automatically pull receipts directly.  
After looking around, I realised there wasn’t a usable public API available, so that approach was dead.

Then I noticed something obvious: every time I buy a ticket, I also receive an email confirmation containing the receipt PDF.

That was the breakthrough.

Instead of relying on an API that didn’t exist, I could simply pull the receipts from my inbox automatically.

I originally considered using the Gmail API directly, but while researching I found **Google Apps Script** at:

https://script.google.com

It turned out to be the perfect solution, because it gives direct access to Gmail and Google Drive using built-in services.

That’s how this project came into being.

---

## What This Script Does

This script automatically:

- Searches Gmail for ticket-related emails from a specific sender
- Filters only the journeys I care about (Leeds <-> Hull)
- Downloads the attached PDF e-receipts
- Saves them into Google Drive
- Organises them neatly into folders by year and month
- Labels emails after processing so nothing is downloaded twice

It creates a Drive structure like this:

Train Tickets /
2026 /
February /
Feb - Sat - 07-02-2026.pdf
Feb - Mon - 10-02-2026.pdf