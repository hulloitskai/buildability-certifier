# Buildability Certifier

The certifier is an electron-based fully-automatic content-encrypted (unmodifiable without password) pdf certificate creator and sender. It loads files from an Excel workbook, and sends out emails in batches. 

It is designed to better help Buildability complete client certification whenever a large class of clients complete a Buildability training course.

![](https://github.com/steven-xie/buildability-certifier/blob/master/image-showcase/ui-main.png)

## Features
- Has a fully-automatic mode and a semi-automatic mode (individual client name selection with autocomplete through Awesomplete).
- Uses SMTP to send emails through any provider (options available for SMTP host, email username & password).
- Beautiful interface designed on Sketch and implemented through CSS.
- Custom options for email per batch
- Has a default certificate background, but contains a file-selector for other certificate backgrounds.
- Uses 256-bit encryption to lockdown PDF-modification through QPDF.
- Custom interface to select the Excel columns that contain the necessary information to fill out the certificate.
- After a set of emails, any email-timeouts due to serverload are recorded; User is prompted if they want to attempt a second send only for the emails that fail.
- Data-preservation; Backs up Excel files in case of file corruption. Cache is set to 20 files at a time (modifiable via settings), with auto-deletion on oldest file when limit is reached. 
- Writes to Excel column, inserts 'true' for clients that were emailed successfully.

### Releases
You can [download the latest Windows release of the app here](https://github.com/steven-xie/buildability-certifier/releases), although it won't do you a whole of good if you're not a employee at [Buildability](http://www.buildability.ca);
