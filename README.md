# Automatic Mail Sender for SMV 2024 Volunteers

## Instructions

1. Place the `openhelpers.xlsx` file (from Helfereinsatz.ch) in the root directory.
2. Run `main.py`.

## Configuration Settings

- `send_email_flag = true`  # Set to `false` for debugging purposes
- `batch_size = 1`          # Number of emails sent per batch
- `pause_duration = 60`     # Pause duration (in seconds) between batches
- `bcc_sender = true`       # Send a BCC to the sender

