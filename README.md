# jurorssummons AKA Civilian Reviewer Lizania Cruz Interactive Printer Installation

This repository contains all the code used for the Civilian Reviewer Installation, the repo is called Jurors Summons because there was an early miscommunication about the name.  There are a couple of pieces of infrastructure and things to be aware of.

### INSTALL & RUN
```
pip install redis
pip install python-docx
python redis_sub.py
```
### TO TEST PRINT OUT
`python formatPrintout.py`

# Key Pieces of Software
### The Web App
The Web App is the client app that participants will interact with on their phones.  It is a basic NodeJs Express app which ran on Heroku.  The key piece that makes this work is a library called fingerprint which is able to create a browser fingerprint.  This would typically be used for ad based tracking but we can use it here to easily know when one participant has moved from one station to another.  When a fingerprint id is generated it gets sent into a Redis DB which is the back end db for the entire system.

### The Redis DB
The Redis DB is the key to the entire project (I used the free tier from redis labs), you should be able to just sign up for your own account and update the DB connection details.  Every time someone scans a QR code and the page loads the web app looks at the fingerprint of the user to see if they have registered or not, if they haven't they will be asked to go back to station 1.  If they are in the DB already the redis DB will return all their previous answers.

### The Raspberry Pi
On The raspberry Pi you need to have a reliable internet connection and have the raspberry pi auto start the file `redis_sub.py`.  That file is listening for the pubsub notification that someone has completed all their questions.  If it passes a basic validation, it will use the code in `formatPrintout.py` to create a word document(using python docx library) out of the answers the participant has placed.

#### Printing
At the end of `formatPrintout.py` you will see a python subprocess call that will tell the printer to print out the newly created document.  That's it!

### Printer Setup
If you are using a new raspberry pi you need to setup the printer using the CUPS Admin page.

# Final Thoughts
I apoligize for the lack of documentation, there are extra other csv/text file datasets and utlity files but everything you need should be in this repo.  feel free to reach out to me if you have any questions.

# Good Luck!
