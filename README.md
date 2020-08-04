Google Sheets - scripts based solution, driven by custom app script to automate velocity and quality data collection and reporting across buildout Sprints.

The first step is to import the chrome add-on “JIRA cloud for sheets” to essentially pull data from JIRA. It supports the following:
- Any JQL needed to pull data from JIRA.
- Ability to select fields to be imported (The default is the columns selected by the user In JIRA)
- Ability to schedule the import – for e.g. once every day, once every hour We import all this data into a tab by the name “StoriesDump.” This name can be anything you can choose,
  but this name is used by the code in steps below.
