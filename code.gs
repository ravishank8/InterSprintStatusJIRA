  /**
   * On open sheet
   */
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('JIRA Menubar')
      .addItem('Update Story Status', 'getStoryCompletionStatus')
      .addItem('Update Features Status', 'getFeatureCompletionStatus')
      .addItem('Features Turning Green Today', 'featuresTurningGreenToday')
      .addToUi();
  }

  //Rounding function
  function roundToTwo(num) {
    return +(Math.round(num + "e+2") + "e-2");
  }

  var devNotStartedStatusCategory = "Dev Not Started";
  var devInProgressStatusCategory = "In Progress";
  var devCompleteStatusCategory = "Dev Complete";
  var storyCompletedStatusCategory = "SP Achieved";

  /**
   * This function is used to log into google sheets story information from JIRA
   *
   */
  function getStoryCompletionStatus() {
    var storiesDumpSheet = SpreadsheetApp.getActive().getSheetByName('StoriesDump');
    var storiesAnalysisSheet = SpreadsheetApp.getActive().getSheetByName('StoriesChart');

    var data = storiesDumpSheet.getDataRange().getValues();

    var finalStoryStatus = [];

    //First get the indexes right for the Actual Start Date, Actual End Date, Planned Start Date and Due Date
    var storyIdIndex = data[0].indexOf("Key");
    var podIndex = data[0].indexOf("RWD POD");
    var assigneeIndex = data[0].indexOf("Assignee");
    var storyPointsIndex = data[0].indexOf("Story Points");
    var plannedEndDateColumnIndex = data[0].indexOf("Due Date");
    var originalEstimateIndex = data[0].indexOf("Σ Original Estimate");
    var remainingTimeIndex = data[0].indexOf("Σ Remaining Estimate");
    var storyStatusIndex = data[0].indexOf("Status");
    var storySprintIndex = data[0].indexOf("Sprint");
    var componentsIndex = data[0].indexOf("Components");

    var today = new Date();

    storiesAnalysisSheet.clear();
    storiesAnalysisSheet.appendRow(["Key", "Responsive POD", "Original Estimate", "Remaining Effort", "Assignee", "Completion Status", "JIRA Status", "Story Points", "Story Points Achieved", "Story Points Developed", "Sprint", "Components", "Due Date"]);

    for (var rownum = 1; rownum < data.length; rownum++) {

      var storyKey = data[rownum][storyIdIndex];
      var storyPoints = data[rownum][storyPointsIndex];
      var storyStatus = data[rownum][storyStatusIndex];
      var estimatedEffort = data[rownum][originalEstimateIndex] / 28800;
      var effortRemaining = data[rownum][remainingTimeIndex] / 28800;
      var jiraStatus = data[rownum][storyStatusIndex];
      var storyStatusCategory = getStoryCompletionCategory(storyStatus);
      var storyPointAchieved = 0;
      var storyPointsDeveloped = 0;
      var storySprint = data[rownum][storySprintIndex];

      if (storyStatusCategory === storyCompletedStatusCategory) {
        storyPointAchieved = storyPoints;
        storyPointsDeveloped = storyPoints;
      } else if (storyStatusCategory === devCompleteStatusCategory) {
        storyPointsDeveloped = storyPoints;
      }
      finalStoryStatus.push([data[rownum][storyIdIndex], data[rownum][podIndex], data[rownum][originalEstimateIndex], data[rownum][remainingTimeIndex], data[rownum][assigneeIndex], storyStatusCategory, jiraStatus, storyPoints, storyPointAchieved, storyPointsDeveloped, storySprint, data[rownum][componentsIndex], data[rownum][plannedEndDateColumnIndex]]);

      // Commented out performance intensive code
      //storiesAnalysisSheet.appendRow([data[rownum][storyIdIndex], data[rownum][podIndex], data[rownum][originalEstimateIndex], data[rownum][remainingTimeIndex], data[rownum][assigneeIndex], storyStatusCategory, jiraStatus, storyPoints, storyPointAchieved, storyPointsDeveloped, storySprint, data[rownum][componentsIndex], data[rownum][plannedEndDateColumnIndex]]);
    }

    // Batch Update
    storiesAnalysisSheet.getRange(storiesAnalysisSheet.getLastRow() + 1, 1, finalStoryStatus.length, finalStoryStatus[0].length).setValues(finalStoryStatus);

    console.log("hello");
  }

  /*
   * Utility function to story status completion
   */
  function getStoryCompletionCategory(storyStatus) {
    var storyCompleteStatusString = "In PO Demo, Ready for PO Demo,  READY FOR UAT";
    var devCompleteStatusString = "READY FOR TESTING, In Test, Testing Failed";
    var incompleteStatusString = "IN DEV, Blocked, ";
    var devNotStartedString = "To Do, In PO Review, READY FOR DEV, Ready for Grooming, IN ANALYSIS, In Grooming, READY FOR PLANNING, IN REFINEMENT";

    if (incompleteStatusString.indexOf(storyStatus) !== -1) {
      return devInProgressStatusCategory;
    } else if (devNotStartedString.indexOf(storyStatus) !== -1) {
      return devNotStartedStatusCategory;
    } else if (devCompleteStatusString.indexOf(storyStatus) !== -1) {
      return devCompleteStatusCategory;
    } else {
      return storyCompletedStatusCategory;
    }
  }

  /*
   * Function to populate in google sheets the feature completion status
   */
  function getFeatureCompletionStatus() {
    var storiesAnalysisSheet = SpreadsheetApp.getActive().getSheetByName('StoriesChart');
    var defectsDumpsheet = SpreadsheetApp.getActive().getSheetByName("Defects Dump");
    var featuresCompletionSheet = SpreadsheetApp.getActive().getSheetByName("Feature-wise Completion");
    var mvpFeaturesSheet = SpreadsheetApp.getActive().getSheetByName("MVP List");

    var storiesData = storiesAnalysisSheet.getDataRange().getValues();
    var defectsData = defectsDumpsheet.getDataRange().getValues();
    var mvpData = mvpFeaturesSheet.getDataRange().getValues();

    var featureIndexInStoryDump = storiesData[0].indexOf("Components");
    var featureIndexInDefectsDump = defectsData[0].indexOf("Components");
    var totalSPIndexInStoryDump = storiesData[0].indexOf("Story Points");
    var SPAchievedIndexInStoryDump = storiesData[0].indexOf("Story Points Achieved");
    var SPDevelopedIndexInStoryDump = storiesData[0].indexOf("Story Points Developed");
    var podDataIndex = storiesData[0].indexOf("Responsive POD");
    var storySprintIndex = storiesData[0].indexOf("Sprint");
    var dueDateIndex = storiesData[0].indexOf("Due Date");
    var storyStatusIndex = storiesData[0].indexOf("JIRA Status");

    var defectKeyIndex = defectsData[0].indexOf("Key");
    var defectPriorityIndex = defectsData[0].indexOf("Priority");
    var defectRankingIndex = defectsData[0].indexOf("Labels");
    var defectCategoryIndex = defectsData[0].indexOf("Category");

    var mvpComponentIndex = mvpData[0].indexOf("Component");
    var isMVPIndex = mvpData[0].indexOf("MVP");

    var featuresData = [];

    for (var rownum = 0; rownum < storiesData.length; rownum++) {
      featuresData.push([storiesData[rownum][podDataIndex], storiesData[rownum][featureIndexInStoryDump]]);
    }

    var unique = function (value, index, self) {
      return self.indexOf(value[1]) === index;
    }

    var uniqueFeaturesList = [];

    for (var rownum = 0; rownum < featuresData.length; rownum++) {
      if (uniqueFeaturesList.length === 0) {
        uniqueFeaturesList.push(featuresData[rownum]);
      } else {
        var matched = false;
        for (var rows = 0; rows < uniqueFeaturesList.length; rows++) {
          if (featuresData[rownum][1] === uniqueFeaturesList[rows][1]) {
            matched = true;
            break;
          }
        }
        if (matched === false) {
          uniqueFeaturesList.push(featuresData[rownum]);
        }
      }
    }

    //Add MVP to features
    for (var i = 0; i < uniqueFeaturesList.length; i++) {
      for (var count = 0; count < mvpData.length; count++) {
        if ((mvpData[count][mvpComponentIndex]).trim() === uniqueFeaturesList[i][1].trim()) {
          uniqueFeaturesList[i].push(mvpData[count][isMVPIndex]);
          console.log("Value of i is:", i);
          break;
        }
      }
    }

    featuresCompletionSheet.clear();
    featuresCompletionSheet.appendRow(["Component", "RWD POD", "isMVP", "Overall Status",
      "Feature Complete Sprint", "Total Story Points",
      "Story Points Achieved", "Story Points Developed",
      "Story Points Failed", "Blocker Count", "Blocker Rank-1 Count",
      "Blocker Rank 2 Count", "Critical Count", "Critical Rank-1 Count",
      "Critical Rank-2 Count", "Major Count", "Major Rank-1 Count",
      "Major Rank-2 Count", "Minor Count", "Minor Rank-1 Count",
      "Minor Rank-2 Count", "Third Party Count", "Third Party Rank-1 Count",
      "Third Party Rank-2 Count", "Blocker Defects", "Blocker Rank-1 Defects",
      "Blocker Rank-2 Defects", "Critical Defects", "Critical Rank-1 Defects",
      "Critical Rank-2 Defects", "Major Defects", "Major Rank-1 Defects",
      "Major Rank-2 Defects", "Minor Defects", "Minor Rank-1 Defects",
      "Minor Rank-2 Defects", "Third Party Defects", "Third Party Rank-1 Defects",
      "Third Party Rank-2 Defects", "Feature End Date", "Feature Ending Week", "Total Defect Count"
    ]);

    var finalFeaturesArray = [];

    for (var rownum = 0; rownum < uniqueFeaturesList.length; rownum++) {

      if ((uniqueFeaturesList[rownum][1] === "Components")) {
        continue;
      }
      var defectsMappedtoFeature = {
        totalDefectCount: 0,
        ThirdParty: {
          Count: 0,
          IDs: "",
          Rank_1_Count: 0,
          Rank_1_IDs: "",
          Rank_2_Count: 0,
          Rank_2_IDs: ""
        },
        Blockers: {
          Count: 0,
          IDs: "",
          Rank_1_Count: 0,
          Rank_1_IDs: "",
          Rank_2_Count: 0,
          Rank_2_IDs: ""
        },
        Criticals: {
          Count: 0,
          IDs: "",
          Rank_1_Count: 0,
          Rank_1_IDs: "",
          Rank_2_Count: 0,
          Rank_2_IDs: ""
        },
        Majors: {
          Count: 0,
          IDs: "",
          Rank_1_Count: 0,
          Rank_1_IDs: "",
          Rank_2_Count: 0,
          Rank_2_IDs: ""
        },
        Minors: {
          Count: 0,
          IDs: "",
          Rank_1_Count: 0,
          Rank_1_IDs: "",
          Rank_2_Count: 0,
          Rank_2_IDs: ""
        },
        storyPoints: {
          Completed: 0,
          Developed: 0,
          Remaining: 0
        }
      }

      // Map Defects to Features
      for (var defectsRow = 0; defectsRow < defectsData.length; defectsRow++) {
        if (defectsData[defectsRow][featureIndexInDefectsDump] === uniqueFeaturesList[rownum][1]) {

          defectsMappedtoFeature.totalDefectCount++;
          // Populate overall Defect count by priority
          if (defectsData[defectsRow][defectPriorityIndex] === "BLOCKER") {
            defectsMappedtoFeature.Blockers.Count += 1;
            defectsMappedtoFeature.Blockers.IDs = defectsMappedtoFeature.Blockers.IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_1") !== -1) {
              defectsMappedtoFeature.Blockers.Rank_1_Count += 1;
              defectsMappedtoFeature.Blockers.Rank_1_IDs = defectsMappedtoFeature.Blockers.Rank_1_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_2") !== -1) {
              defectsMappedtoFeature.Blockers.Rank_2_Count += 1;
              defectsMappedtoFeature.Blockers.Rank_2_IDs = defectsMappedtoFeature.Blockers.Rank_2_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }

          } else if (defectsData[defectsRow][defectPriorityIndex] === "CRITICAL") {
            defectsMappedtoFeature.Criticals.Count += 1;
            defectsMappedtoFeature.Criticals.IDs = defectsMappedtoFeature.Criticals.IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_1") !== -1) {
              defectsMappedtoFeature.Criticals.Rank_1_Count += 1;
              defectsMappedtoFeature.Criticals.Rank_1_IDs = defectsMappedtoFeature.Criticals.Rank_1_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_2") !== -1) {
              defectsMappedtoFeature.Criticals.Rank_2_Count += 1;
              defectsMappedtoFeature.Criticals.Rank_2_IDs = defectsMappedtoFeature.Criticals.Rank_2_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }

          } else if (defectsData[defectsRow][defectPriorityIndex] === "MAJOR") {
            defectsMappedtoFeature.Majors.Count += 1;
            defectsMappedtoFeature.Majors.IDs = defectsMappedtoFeature.Majors.IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_1") !== -1) {
              defectsMappedtoFeature.Majors.Rank_1_Count += 1;
              defectsMappedtoFeature.Majors.Rank_1_IDs = defectsMappedtoFeature.Majors.Rank_1_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_2") !== -1) {
              defectsMappedtoFeature.Majors.Rank_2_Count += 1;
              defectsMappedtoFeature.Majors.Rank_2_IDs = defectsMappedtoFeature.Majors.Rank_2_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }

          } else if (defectsData[defectsRow][defectPriorityIndex] === "MINOR") {
            defectsMappedtoFeature.Minors.Count += 1;
            defectsMappedtoFeature.Minors.IDs = defectsMappedtoFeature.Minors.IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_1") !== -1) {
              defectsMappedtoFeature.Minors.Rank_1_Count += 1;
              defectsMappedtoFeature.Minors.Rank_1_IDs = defectsMappedtoFeature.Minors.Rank_1_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_2") !== -1) {
              defectsMappedtoFeature.Minors.Rank_2_Count += 1;
              defectsMappedtoFeature.Minors.Rank_2_IDs = defectsMappedtoFeature.Minors.Rank_2_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }
          }

          if (defectsData[defectsRow][defectCategoryIndex] === "Third Party") {
            defectsMappedtoFeature.ThirdParty.Count += 1;
            defectsMappedtoFeature.ThirdParty.IDs = defectsMappedtoFeature.ThirdParty.IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_1") !== -1) {
              defectsMappedtoFeature.ThirdParty.Rank_1_Count += 1;
              defectsMappedtoFeature.ThirdParty.Rank_1_IDs = defectsMappedtoFeature.ThirdParty.Rank_1_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }

            if ((defectsData[defectsRow][defectRankingIndex]).toUpperCase().indexOf("RANK_2") !== -1) {
              defectsMappedtoFeature.ThirdParty.Rank_2_Count += 1;
              defectsMappedtoFeature.ThirdParty.Rank_2_IDs = defectsMappedtoFeature.ThirdParty.Rank_2_IDs.concat(defectsData[defectsRow][defectKeyIndex]).concat(";");
            }
          }
        }
      }

      // Map stories to features
      var totalStoryPoints = 0;
      var storyPointsAchieved = 0;
      var storyPointsDeveloped = 0;
      var featureEndingSprint = "";
      var featureEndingDateArray = [];
      var totalStoryPointsFailed = 0;
      var storyStatusIndexConcat = ";";

      for (var storiesRow = 0; storiesRow < storiesData.length; storiesRow++) {
        if (storiesData[storiesRow][featureIndexInStoryDump] === uniqueFeaturesList[rownum][1]) {
          featureEndingDateArray.push(new Date(storiesData[storiesRow][dueDateIndex]));
          if (storiesData[storiesRow][totalSPIndexInStoryDump] !== "") {
            featureEndingSprint = featureEndingSprint.concat(storiesData[storiesRow][storySprintIndex]);
            totalStoryPoints += storiesData[storiesRow][totalSPIndexInStoryDump];
          }
          if (storiesData[storiesRow][SPAchievedIndexInStoryDump] !== "") {
            storyPointsAchieved += storiesData[storiesRow][SPAchievedIndexInStoryDump];
          }
          if (storiesData[storiesRow][SPDevelopedIndexInStoryDump] !== "") {
            storyPointsDeveloped += storiesData[storiesRow][SPDevelopedIndexInStoryDump];
          }

          storyStatusIndexConcat = storyStatusIndexConcat.concat(storiesData[storiesRow][storyStatusIndex]).concat(";");

          // Handle Testing Failed Status
          if (storiesData[storiesRow][storyStatusIndex] === "Testing Failed") {
            totalStoryPointsFailed += storiesData[storiesRow][totalSPIndexInStoryDump];
          }
        }
      }

      // Get Status Data
      var overallStatus = "";
      var featureEndDateObj = "";
      var featureEndDateStr = "";
      var featureEndDate = "";
      var featureEndWeek = "";

      if ((uniqueFeaturesList[rownum][1] === "") || (uniqueFeaturesList[rownum][1] === "NFR") || (uniqueFeaturesList[rownum][1].indexOf("Tech Enabler") !== -1)) {
        overallStatus = "NA";
        featureEndingSprint = "NA";
      } else {
        // Not started Status
        if (storyPointsDeveloped === 0) {
          if ((storyStatusIndexConcat.toUpperCase().indexOf("IN DEV")) !== -1) {
            overallStatus = "Dev In Progress";
          } else {
            overallStatus = "Not Started";
          }

          featureEndDateObj = getFeatureEndDate(featureEndingDateArray);
          featureEndDate = featureEndDateObj.featureEndDate;
          featureEndDateStr = featureEndDateObj.featureEndDateStr;
          featureEndWeek = featureEndDateObj.featureEndDateWeekEnding;
        }
        // In Progress Status
        if ((storyPointsDeveloped !== 0) && (storyPointsAchieved < totalStoryPoints)) {
          if (storyPointsDeveloped === totalStoryPoints) {
            overallStatus = "Testing in Progress";
            featureEndingSprint = "Completed";
          }
          if (storyPointsDeveloped < totalStoryPoints) {
            overallStatus = "Dev In Progress";
            if (featureEndingSprint.indexOf("Sprint 12") !== -1) {
              featureEndingSprint = "Sprint 12";
            } else if (featureEndingSprint.indexOf("Sprint 11") !== -1) {
              featureEndingSprint = "Sprint 11";
            } else if (featureEndingSprint.indexOf("Sprint 10") !== -1) {
              featureEndingSprint = "Sprint 10";
            } else {
              featureEndingSprint = "TBD";
            }
          }
          featureEndDateObj = getFeatureEndDate(featureEndingDateArray);
          featureEndDate = featureEndDateObj.featureEndDate;
          featureEndDateStr = featureEndDateObj.featureEndDateStr;
          featureEndWeek = featureEndDateObj.featureEndDateWeekEnding;
        }

        // Complete with defects
        // Added Ranking Logic - Really!
        if ((storyPointsDeveloped !== 0) && (storyPointsAchieved === totalStoryPoints)) {
          // if ((defectsMappedtoFeature.Blockers.Count > 0) ||
          //   (defectsMappedtoFeature.Criticals.Count > 0) ||
          //   (defectsMappedtoFeature.Majors.Count > 4)) {

          if ((defectsMappedtoFeature.Blockers.Count > 0) ||
            (defectsMappedtoFeature.Criticals.Count > 0) ||
            (defectsMappedtoFeature.Majors.Rank_1_Count > 2)) {
            overallStatus = "Complete With Defects";
            featureEndingSprint = "Completed";
          } else {
            overallStatus = "Ready For UAT";
            featureEndingSprint = "Completed";
          }
        }
      }
      // Commented out performance intensive code
      //featuresCompletionSheet.appendRow([uniqueFeaturesList[rownum][1], uniqueFeaturesList[rownum][0],uniqueFeaturesList[rownum][2], overallStatus, featureEndingSprint, totalStoryPoints, storyPointsAchieved, storyPointsDeveloped, totalStoryPointsFailed, defectsMappedtoFeature.Blockers.Count, defectsMappedtoFeature.Criticals.Count, defectsMappedtoFeature.Majors.Count, defectsMappedtoFeature.Minors.Count, defectsMappedtoFeature.Blockers.IDs, defectsMappedtoFeature.Criticals.IDs, defectsMappedtoFeature.Majors.IDs, defectsMappedtoFeature.Minors.IDs, featureEndDate, featureEndWeek]);
      finalFeaturesArray.push([uniqueFeaturesList[rownum][1], uniqueFeaturesList[rownum][0],
        uniqueFeaturesList[rownum][2], overallStatus, featureEndingSprint, totalStoryPoints,
        storyPointsAchieved, storyPointsDeveloped, totalStoryPointsFailed,
        defectsMappedtoFeature.Blockers.Count, defectsMappedtoFeature.Blockers.Rank_1_Count,
        defectsMappedtoFeature.Blockers.Rank_2_Count, defectsMappedtoFeature.Criticals.Count,
        defectsMappedtoFeature.Criticals.Rank_1_Count, defectsMappedtoFeature.Criticals.Rank_2_Count,
        defectsMappedtoFeature.Majors.Count, defectsMappedtoFeature.Majors.Rank_1_Count,
        defectsMappedtoFeature.Majors.Rank_2_Count, defectsMappedtoFeature.Minors.Count,
        defectsMappedtoFeature.Minors.Rank_1_Count, defectsMappedtoFeature.Minors.Rank_2_Count,
        defectsMappedtoFeature.ThirdParty.Count, defectsMappedtoFeature.ThirdParty.Rank_1_Count,
        defectsMappedtoFeature.ThirdParty.Rank_2_Count, defectsMappedtoFeature.Blockers.IDs, 
        defectsMappedtoFeature.Blockers.Rank_1_IDs, defectsMappedtoFeature.Blockers.Rank_2_IDs, 
        defectsMappedtoFeature.Criticals.IDs, defectsMappedtoFeature.Criticals.Rank_1_IDs, 
        defectsMappedtoFeature.Criticals.Rank_2_IDs, defectsMappedtoFeature.Majors.IDs, 
        defectsMappedtoFeature.Majors.Rank_1_IDs, defectsMappedtoFeature.Majors.Rank_2_IDs, 
        defectsMappedtoFeature.Minors.IDs, defectsMappedtoFeature.Minors.Rank_1_IDs, 
        defectsMappedtoFeature.Minors.Rank_2_IDs, defectsMappedtoFeature.ThirdParty.IDs, 
        defectsMappedtoFeature.ThirdParty.Rank_1_IDs, defectsMappedtoFeature.ThirdParty.Rank_2_IDs,
        featureEndDate, featureEndWeek, defectsMappedtoFeature.totalDefectCount
      ]);
    }

    // Batch Update
    featuresCompletionSheet.getRange(featuresCompletionSheet.getLastRow() + 1, 1, finalFeaturesArray.length, finalFeaturesArray[0].length).setValues(finalFeaturesArray);
  }

  // Utility function for feature completion
  function getFeatureEndDate(featureEndingDateArray) {
    var featureEndDateObj = {
      featureEndDate: "",
      featureEndDateStr: "",
      featureEndDateWeekEnding: ""
    };

    var featureEndDate = new Date('1/1/2019');
    var featureEndDateStr = "";
    var featureEndDateWeekEnding = new Date('1/1/2019');

    for (var i = 0; i < featureEndingDateArray.length; i++) {
      featureEndDateStr = featureEndDateStr.concat(featureEndingDateArray[i]).concat(";");
      if (featureEndingDateArray[i] !== "") {
        if (featureEndingDateArray[i] > featureEndDate) {
          featureEndDate = featureEndingDateArray[i];
        }
      }
    }
    featureEndDate.setDate(featureEndDate.getDate() + 2);
    featureEndDateObj.featureEndDate = featureEndDate;
    featureEndDateObj.featureEndDateStr = featureEndDateStr;

    featureEndDateWeekEnding = new Date(featureEndDate.getTime());

    if (featureEndDateWeekEnding.getDay() < 6) {
      featureEndDateWeekEnding.setDate(featureEndDateWeekEnding.getDate() + (6 - featureEndDateWeekEnding.getDay()));
    }
    featureEndDateObj.featureEndDateWeekEnding = featureEndDateWeekEnding;

    return featureEndDateObj;
  }

  // How do defects assigned today impact features - Runtime data
  function featuresTurningGreenToday() {
    var featureDataSheet = SpreadsheetApp.getActive().getSheetByName('Feature-wise Completion');
    var defectForTodayDataSheet = SpreadsheetApp.getActive().getSheetByName("Defects Assigned Today");
    var defectsFeatureCompletionSheet = SpreadsheetApp.getActive().getSheetByName("Defects-Feature Movement");

    var featureData = featureDataSheet.getDataRange().getValues();
    var defectForTodayData = defectForTodayDataSheet.getDataRange().getValues();

    var featureNameIndex = featureData[0].indexOf("Component");
    var featureStatusIndex = featureData[0].indexOf("Overall Status");
    var blockerDefectIndex = featureData[0].indexOf("Blocker Count");
    var criticalDefectIndex = featureData[0].indexOf("Critical Count");
    var majorRank1DefectIndex = featureData[0].indexOf("Major Rank-1 Count");

    var inTodayDefectsComponentIndex = defectForTodayData[0].indexOf("Components");
    var inTodayDefectsPriorityIndex = defectForTodayData[0].indexOf("Priority");
    var inTodayDefectsLabelsIndex = defectForTodayData[0].indexOf("Labels");

    var blockerDefectCount = 0;
    var criticalDefectCount = 0;
    var majorRank1DefectCount = 0;
    var featureName = "";

    var featuresToBeTurnedGreen = [];
    var featuresToPrioritizeTesting = [];

    for (var rownum = 1; rownum < featureData.length; rownum++) {
      if (((featureData[rownum][featureStatusIndex]) === "Ready For UAT") ||
        ((featureData[rownum][featureStatusIndex]) === "Dev In Progress") ||
        ((featureData[rownum][featureStatusIndex]) === "NA") ||
        ((featureData[rownum][featureStatusIndex]) === "Not Started")) {
        continue;
      } else {
        featureName = featureData[rownum][featureNameIndex];
        blockerDefectCount = featureData[rownum][blockerDefectIndex];
        criticalDefectCount = featureData[rownum][criticalDefectIndex];
        majorRank1DefectCount = featureData[rownum][majorRank1DefectIndex];
        for (var defectNum = 1; defectNum < defectForTodayData.length; defectNum++) {
          if (defectForTodayData[defectNum][inTodayDefectsComponentIndex] === featureName) {
            if (defectForTodayData[defectNum][inTodayDefectsPriorityIndex] === "BLOCKER") {
              blockerDefectCount--;
            } else if (defectForTodayData[defectNum][inTodayDefectsPriorityIndex] === "CRITICAL") {
                criticalDefectCount--;
            } else if (defectForTodayData[defectNum][inTodayDefectsPriorityIndex] === "MAJOR") {
              if ((defectForTodayData[defectNum][inTodayDefectsLabelsIndex].toUpperCase().indexOf("RANK_1") !== -1)) {
                majorRank1DefectCount--;
              }
            }
          }
        }
        if ((blockerDefectCount === 0) && (criticalDefectCount === 0) && (majorRank1DefectCount < 2)) {
          if ((featureData[rownum][featureStatusIndex]) === "Testing in Progress") {
            featuresToPrioritizeTesting.push(featureName);
          } else {
            featuresToBeTurnedGreen.push(featureName);
          }
        }
      }
    }
   // console.log("Completed");
    defectsFeatureCompletionSheet.clear();
    defectsFeatureCompletionSheet.appendRow(["FEATURES - POTENTIALLY TURNING GREEN"]);
    for (var i = 0; i < featuresToBeTurnedGreen.length; i++) {
      defectsFeatureCompletionSheet.appendRow([featuresToBeTurnedGreen[i]]);
    }

    defectsFeatureCompletionSheet.appendRow(["FEATURES IN TESTING - TO PRIORITIZE FOR TESTING"]);
    for (var i = 0; i < featuresToPrioritizeTesting.length; i++) {
      defectsFeatureCompletionSheet.appendRow([featuresToPrioritizeTesting[i]]);
    }
  }