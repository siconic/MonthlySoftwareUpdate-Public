# MonthlySoftwareUpdate-Public

MECM Monthly Software Update Script

This script will automate ALL functions done during patching. It creates a folder structure for the updates, renames the Deployment package 
in the ADR, Runs the ADR, Renames the Software Update Group, creates the deployments on a schedule, has customizable parameters for the 
deployments, and emails the deployed patches and scheduled deployment times to whatever distribution list or list of users you wish.

All aspects of the deployments are customizable. There is a time code bug, that shifts the time by about an hour, sometimes 30 minutes. 
I am not sure what causes this, and it seems to happen 2 or 3 times per year. Due to this bug in the time code, I suggest you always use 
the option "-Enable $false" and manually enable the deployments once the times are confirmed. If the exact times are not important to you, 
then you can disregard this. 

Requires an ADR to function. The ADR CAN NOT be set to run automatically. The name of the ADR is a variable in the script, but I recomend 
"ADR for Monthly Security Updates".

This script COULD be setup to run as a Scheduled task, to FULLY automate the process. I have personally never done this. 

Pre-Requisits: An ADR will need to be created. There are a set of REQUIRED settings and the software Updates Section is highly customizable
       and can be whatever you want. I suggest you use the following;

The following settings in the ADR ARE REQUIRED;
1. Deployment Settings - Deploy and Approve
2. Do not run this rule automatically
3. Alerts - Generate alerts when rule fails
4. Deployment Package - Select A deployment Package 

Recomended;
1. Software Updates Date Released
   - Last 4 days
   - Required - >=1
   - Superseded - No
   - Update Classifiction - Critical Updates or Security Updates
3. ADR Name - ADR for Monthly Security Updates 
