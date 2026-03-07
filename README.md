# win-iis-lab
Scripts to set up a few websites on IIS for testing/lab env

Install powershell 7 and .NET 4.8 before running the scripts. Reboot if necessary.

Navigate to C:\LabSetup

From an admin powershell 7 session run the scripts in this order:
- Download-Prerequisites.ps1
- Install-Prerequisites.ps1
- Install-Wordpress.ps1
- Install-JuiceShop.ps1
- Install-EchoSPA.ps1

Scripts will install self-signed cert for HTTPS. All sites will have browser warning but will work if you "accept risk and continue."

- Wordpress: https://wordpress.lab.local (first login to perform WP setup)
- JuiceShop: https://juiceship.lab.local - basic OWASP JuiceShop site for testing
- EchoSPA: https://echo.lab.local - simple single page application that displays HTTP request information

If testing with a WAF, make sure that the connection to the origin server uses these settings:
- HTTPS on port 443
- SNI enabled
- TLS v1.2 and/or v1.3 enabled

  
