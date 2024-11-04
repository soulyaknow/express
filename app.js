const Service = require("node-windows").Service;
const path = require("path");

const svc = new Service({
  name: "TagUIService", // Name of the Windows Service
  description: "Node.js web service running as a Windows service",
  script:
    "C:/Users/Hello World!/AppData/Local/ai-broker-service/express/app.js", // Path to app.js
  nodeOptions: ["--harmony", "--max_old_space_size=4096"],
});

if (process.argv.includes("--install")) {
  svc.on("install", () => {
    console.log("Service installed successfully!");
    svc.start();
  });
  svc.install();
} else if (process.argv.includes("--uninstall")) {
  svc.on("uninstall", () => {
    console.log("Service uninstalled successfully!");
  });
  svc.uninstall();
} else {
  // Start the Express web service if no install/uninstall argument is passed
  const express = require("express");
  const cors = require("cors");
  const bodyParser = require("body-parser");
  const { execSync } = require("child_process");
  const { exec } = require("child_process");
  const fs = require("fs");
  const path = require("path");

  const app = express();
  app.use(bodyParser.json());
  app.use(cors());

  const GITLAB_API_URL =
    "https://gitlab.com/api/v4/projects/YOUR_PROJECT_ID/repository/commits";
  const GITLAB_TOKEN = "YOUR_GITLAB_PERSONAL_ACCESS_TOKEN";

  let latestCommitHash = null;

  // Function to check for updates from GitLab before starting the app
  async function checkForUpdates() {
    try {
      const response = await fetch(`${GITLAB_API_URL}?ref_name=main`, {
        headers: {
          "PRIVATE-TOKEN": GITLAB_TOKEN,
        },
      });
      if (!response.ok) throw new Error("GitLab repository not reachable.");

      const data = await response.json();
      const latestCommit = data[0]?.id;

      // Check if there's a new commit
      if (!latestCommitHash || latestCommit !== latestCommitHash) {
        console.log("New update found, pulling changes...");
        latestCommitHash = latestCommit; // Update stored hash after pulling changes
        execSync("git pull origin main"); // Pull latest changes from the repository
      } else {
        console.log("No updates found, starting application normally.");
      }
    } catch (error) {
      console.error(`Error checking for updates: ${error.message}`);
      console.log("Starting application normally...");
    }
  }

  // Endpoint to execute the TagUI script with dynamic data
  app.post("/execute-tagUI-script", (req, res) => {
    const {
      applicationData: {
        "Housing Expense": housingExpense,
        "Household Liabilities": householdLiabilities,
        License: license,
        Passport: passport,
        "Insurance Expense": insuranceExpense,
        Dependents: dependents,
        "Healthcare Expense": healthcareExpense,
        "Household Income": householdIncome,
        "Household Savings": householdSavings,
        Applicants: applicants,
        "Fact Find": factFind,
        Broker: brokerIds,
        "Estimated Settlement Date": estimatedSettlementDate,
        "Loan Type": loanType,
        "Finance Due Date": financeDueDate,
        "App ID": appId,
        Status: status,
        "Education Expense": educationExpense,
        "First Time Home Buyer": firstTimeHomeBuyer,
        "Total Household Expenses": totalHouseholdExpenses,
      },
      brokerData: { recordId, thirdPartyAggregator, thirdPartyCRM },
    } = req.body;

    // The TagUI script
    const taguiScript = `
      // This script is for filling all the necessary form
      // Navigate to test site
      echo "Navigating to test site"
      https://dummycrm.korunaassist.com/
  
      wait for //*[@ng-model="$mdAutocompleteCtrl.scope.searchText" and @aria-label="First name"]
      echo "Filling firstname"
      type //*[@ng-model="$mdAutocompleteCtrl.scope.searchText" and @aria-label="First name"] as ${housingExpense}
  
      wait for //*[@ng-model="$ctrl.contact.familyName"]
      echo "Filling surname"
      type //*[@ng-model="$ctrl.contact.familyName"] as ${householdIncome}
  
      wait for //*[@ng-model="$ctrl.contact.phone"]
      echo "Filling contact number"
      type //*[@ng-model="$ctrl.contact.phone"] as ${loanType}
  
      wait for //*[@ng-model="$ctrl.contact.email"]
      echo "Filling email"
      type //*[@ng-model="$ctrl.contact.email"] as ${appId}
  
      wait for //*[ng-if="!floatingLabel"]
      echo "Filling add team member"
      type //*[@ng-if="!floatingLabel"] as ${estimatedSettlementDate}
  
      wait for //*[@ng-click="next(Model.activeSection)"]
      echo "Clicking next button"
      click //*[@ng-click="next(Model.activeSection)"]
  
      wait 5
  
      wait for //*[@ng-click="$ctrl.ticketLoanSecuritySplitAdd($event)"]
      echo "Clicking add security details button"
      click //*[@ng-click="$ctrl.ticketLoanSecuritySplitAdd($event)"]
  
      wait 5
  
      wait for //*[ng-model="$mdAutocompleteCtrl.scope.searchText"]
      echo "Filling security address"
      type //*[@ng-model="$mdAutocompleteCtrl.scope.searchText"] as ${brokerIds}

      wait 10
  
      echo "Fill process completed"
    `;

    // Define the script directory and file paths
    const scriptDirectory = "C:\\TagUIScripts";
    // const scriptDirectory = "C:\\Users\\Hello World!\\ai-broker";
    const scriptPath = path.join(scriptDirectory, "documentForm.tag");

    // Verify if the script directory exists, create if not
    if (!fs.existsSync(scriptDirectory)) {
      fs.mkdirSync(scriptDirectory, { recursive: true });
    }

    // Write the TagUI script to a file
    try {
      fs.writeFileSync(scriptPath, taguiScript);
      console.log(`TagUI script written to ${scriptPath}`);
    } catch (fileError) {
      console.error(
        `Failed to write TagUI script to file: ${fileError.message}`
      );
      return res.status(500).json({ error: "Failed to write TagUI script" });
    }

    // Use PowerShell to run the command with proper escaping
    const command = `powershell -Command "& { Start-Process 'C:\\TagUI\\src\\tagui.cmd' -ArgumentList \\"${scriptPath}\\", '-t' -Verb RunAs }"`;
    // const command = `powershell -Command "& { Start-Process 'C:\\TagUI\\src\\tagui.cmd' -ArgumentList \\"${scriptPath}\\" -Verb RunAs }" >> C:\\TagUI\\Logs\\tagui_service_log.txt 2>&1`;

    // Execute the command
    exec(command, { cwd: "C:\\TagUI\\src" }, (error, stdout, stderr) => {
      if (error) {
        console.error(`Error executing TagUI script: ${error.message}`);
        return res.status(500).json({
          error: "Failed to execute TagUI script",
          details: error.message,
        });
      }
      if (stderr) {
        console.error(`TagUI script stderr: ${stderr}`);
        return res.status(500).json({
          error: "Error in TagUI script execution",
          details: stderr,
        });
      }

      console.log(`TagUI script executed successfully: ${stdout}`);
      res.json({
        message: "TagUI script executed successfully",
        output: stdout,
      });
    });
  });

  // Default route
  app.get("/", (req, res) => {
    res.send("The TagUI service is running!");
  });

  const PORT = process.env.PORT || 5213;

  // Main function to start the application after checking for updates
  async function startApplication() {
    await checkForUpdates();

    // Start the server
    app.listen(PORT, () => {
      console.log(`Server running on http://localhost:${PORT}`);
    });
  }

  // Run the application
  startApplication();
}
