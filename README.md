# ReceiptTracker
## A gs script to track receipts from a google webapp

Saves data sent to webapp into transaction log in google sheets and uses that data to show monthly expenses  
Saves photos to provided folder ID

Built to be called by Apple Shortcut takes a POST request to webapp address with JSON parameters  
Params:  
apiKey : text  
date : yyyy-mm-dd  
category : text (must match preset options)  
amount : text (sanitizes input for common currency symbols)  
photoBase64 : text optional (photo converted to base64)  



