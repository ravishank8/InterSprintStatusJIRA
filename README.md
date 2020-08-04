Google Sheets - scripts based solution, driven by custom app script to automate velocity and quality data collection and reporting across buildout Sprints.

The first step is to import the chrome add-on “JIRA cloud for sheets” to essentially pull data from JIRA. It supports the following:
- Any JQL needed to pull data from JIRA.
- Ability to select fields to be imported (The default is the columns selected by the user In JIRA)
- Ability to schedule the import – for e.g. once every day, once every hour We import all this data into a tab by the name “StoriesDump.” This name can be anything you can choose,
  but this name is used by the code in steps below.

There are two key tabs in this sheet - 

1- Stories Dump - In this tab we pull in all the stories from JIRA
2- Defects Dump - In this tab we pull in all the defects from JIRA

This is essentially all the information that we need to pull from JIRA. Rest of the work is all about data massaging/merging/operations to publish views for relevant stakeholders - 

In the code.gs: 

- "getStoryCompletionStatus" function is used to populate the storieschart tab, which in turn can drive a pivot to provide succinct story points' status across Sprints

- "getFeatureCompletionStatus" is used to calculate the feature completion status based on stories' completion and defects open against the feature. The same is used to populate the Feature-Wise Completion tab, and the PIVOT corresponding to the same. Notice the various statuses, also the dev-end dates/weeks for features and MVP segregation.

- "featuresTurningGreenToday" is used to project the features projected to turn green today based on defect assignments to developers, and the ones that need to be the focus on quality assurance to test based on the ones assigned to them. This uses an additional data retrieval tab, "Defects assigned today", which is driven by the JQL to get defects assigned to developers today

We also have other tabs that were used to map developers' efficiency with story points, defect projections that were displayed in team area, daily defect target and defects density. But essentially all of these were derivations from the stories' and defects' data that were mapped to features/components.  
