# IMAP Copy
Copy emails and folders from an IMAP account to another one.

Creates missing folders and skips existing messages (using message-id).

Source IMAP is always accessed READ-ONLY.

```
Usage: imapcp.py <user>:<password>:<host>:<port> <user>:<password>:<host>:<port>

Options:
  --version             show program's version number and exit
  -h, --help            show this help message and exit
  -e EXCLUDE, --exclude=EXCLUDE
                        Exclude folders matching pattern (can be specified
                        multiple times)
  -f FOLDER, --folder=FOLDER
                        Only copy a single folder (use from:to to specify a
                        different destinatin name)
  -s, --simulate        Do not perform any task
  --from=FR             Only copy messages older than this date (inclusive)
  --to=TO               Only copy messages newer than this date (inclusive)
  ```
  
