# EduHub Export to OnDemand 

This script will export data for existing, leaving, and future students from EduHub into a valid (see below) OnDemand CSV Student Import file.

## Getting Started
### Prerequisites:
* Script must be run on the same server that hosts eduHub, under an account with at least read access to the local eduHub directory (e.g. D:\eduhub).

### Setup:
1. Download both `EduHub-OnDemand.bat` and `EduHub-OnDemand.ps1` to the admin server that is running EduHub.
2. Modify `EduHub-OnDemand.bat` to include the network share that you wish to export the On Demand data to (e.g. a curric share). **NOTE:** for security reasons, use an account that only has access to that share and nothing else.
3. Create a scheduled task to run `EduHub-OnDemand.bat` under any service account that has read access to the EduHub folder (such as the "Network Service" account). For security do not use the "System" account.
4. Run the Scheduled task and check that `OnDemand.csv` is created in the share that you specified.

## Notes:
* This script is based on Graeme Shea's UserCreator script, as as such can be run together with it; Just add running `EduHub-OnDemand.ps1` to the UserCreator export, and you're good to go.
* If you don't wish to save to a network share, change `$DriveLetter = "L:\"` in `EduHub-OnDemand.ps1` to whatever local folder you want (e.g. `"D:\OnDemand\"`).
* If using the "Network Service" account, make sure that the "Network Service" account *actually* has read access to the eduHub directory. If not, it'll fail to read the data.

## Authors:

* **Graeme Shea** - *Initial Script for UserCreator*
* **Robert Brandon** - *Modified for On Demand*