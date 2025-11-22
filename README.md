# Zendesk ↔ Salla Order History Sync

Automation that syncs **Zendesk private internal notes** into the corresponding **Salla order timeline**, keeping customer order history unified across both systems.

## Problem

- Support agents leave important internal notes inside Zendesk only.
- Salla order history stays empty or outdated.
- Ops / Fulfillment teams working inside Salla cannot see senstive customer interactions with CS team such as:
  - Ticket ID
  - Agent name
  - Timestamp
  - The actual internal note
- This caused repeated questions, missing context, and operational mistakes.

## Solution

A Google Apps Script webhook that:

1. Listens to internal comments submitted inside Zendesk  
2. Extracts the **Salla order number** (21xxxxxx) from the note  
3. Formats a clean Salla note block:
```
#<ticket_id> | <agent_name> | <date>
<cs_note_body>
```
4. Authenticates with Salla using:
  - Client ID
  - Client Secret
  - Refresh token  
5. Appends the note to **Salla order history** via the Merchant API  
6. Ensures API tokens refresh automatically and remain valid

## Tech Stack

- Google Apps Script  
- Zendesk Ticket Audits API  
- Zendesk Users API  
- Salla Merchant API (OAuth2)  

## Key Features

- Syncs Zendesk internal notes → Salla order history  
- Automatically detects Salla order numbers (`21\d{6,}`)  
- Supports multi-store tokens through Script Properties  
- Clean, standardized formatting for Ops visibility  
- Secure webhook using shared key authentication  
- Full OAuth token refresh cycle for Salla

## How It Works (Flow)

1. Zendesk triggers the webhook when an internal note is added  
2. Script:
  - Validates webhook using shared key
  - - Fetches note + agent
    - - Detects Salla order ID  
3. Script calls Salla API:
  - Refresh token if expired
  - - Append formatted note to order history  
4. Returns:
  - `SUCCESS` or `FAILED` + error reason (missing order, invalid token, etc.)

## Configuration

### Script Properties:

  - `SALLA_TOKEN` or `SALLA_TOKENS_JSON`  
  - `ZD_SUBDOMAIN`  
  - `ZD_EMAIL`  
  - `ZD_API_TOKEN`
  - `ORDER_REGEX`  
  - `SHARED_KEY`  
  - Optional: `TZ` (default: Asia/Riyadh)

## Business Impact (Estimate)

- **60–80% reduction** in back-and-forth between CS and Ops  
- Eliminates missing history inside Salla  
- Creates a unified “source of truth” for all order notes  
- Faster issue resolution since Ops sees everything instantly  

## My Role

As **Tech PM & Automation Architect**, I designed the cross-system workflow, implemented the API logic, secured authentication, and standardized the format used across all departments.

