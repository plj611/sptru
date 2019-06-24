### Project TRU

This implementation uses the Sharepoint client object model to manipulate the Sharepoint list. The project has the client and server side. The client is used by user to upload data into the Sharepoint while the server is run as a job to retrieve the list data and generate a file which will upload into a ftp server.

The directory structure for the server is

├───error
│   ├───veirify_ack
│   └───verification
├───notprocessed
│   ├───verification
│   └───verify_ack
├───processed
│   ├───verification
│   └───verify_ack
└───temp
    └───verify_ack