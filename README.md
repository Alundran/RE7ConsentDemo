# Raiser's Edge .NET Consent API Demo

A quick demo Visual Basic forms application for interacting with the COM based Consent API for The Raiser's Edge.

# How to use

  1) Make sure you have The Raiser's Edge installed on your workstation with preferably the latest patch or at least Patch 8
  2) Open the solution in Visual Studio
  3) Edit the Form1.vb code and replace the following fields with your own data:
    -- oApi.Init should have your own database serial number & credentials along with the database number
    -- constit.Load() should contain the System Record ID of the constituent you're adding the consent record to
    -- Consent.Fields should contain your consent channel/category combination specific to your database
   4) Run the application and click on the button
   5) The consent record should be added to the System Record ID you defined
    
This solution code is based on the available API documentation here: https://docs.blackbaud.com/communication-preferences/products/re7/api



