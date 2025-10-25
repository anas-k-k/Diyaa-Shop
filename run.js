const path = require("path");
const fs = require("fs");
const { exec } = require("child_process");

// Set browsers path to our browsers directory
process.env.PLAYWRIGHT_BROWSERS_PATH = path.join(process.cwd(), "browsers");
console.log(`Using browsers from: ${process.env.PLAYWRIGHT_BROWSERS_PATH}`);

// Create a logs directory for error logs
const logsDir = path.join(process.cwd(), "logs");
if (!fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir, { recursive: true });
}

// Create a timestamp for log files
const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
const logFile = path.join(logsDir, `run-log-${timestamp}.txt`);

// Log function to both console and log file
function log(message) {
  const timePrefix = `[${new Date().toISOString()}] `;
  console.log(timePrefix + message);
  try {
    fs.appendFileSync(logFile, timePrefix + message + "\n");
  } catch (e) {
    console.error("Failed to write to log file:", e);
  }
}

log("Starting Diyaa automation test in headed mode...");

let child;
try {
  // Directly run Playwright test via CLI
  const command = `node ./node_modules/@playwright/test/cli.js test tests/login.spec.js --headed`;
  log(`Executing command: ${command}`);

  child = exec(command, {
    env: { ...process.env },
  });

  // Capture and log stdout
  child.stdout.on("data", (data) => {
    log(`STDOUT: ${data.toString().trim()}`);
  });

  // Capture and log stderr
  child.stderr.on("data", (data) => {
    log(`ERROR: ${data.toString().trim()}`);
  });

  // Handle process completion
  child.on("close", (code) => {
    if (code === 0) {
      log("Test completed successfully!");
    } else {
      log(`Test failed with exit code: ${code}`);
      // Check for test result file
      const testResultFile = path.join(
        process.cwd(),
        "test-results",
        ".last-run.json"
      );
      if (fs.existsSync(testResultFile)) {
        try {
          const results = JSON.parse(fs.readFileSync(testResultFile, "utf-8"));
          log(`Test results: ${JSON.stringify(results, null, 2)}`);
        } catch (e) {
          log(`Failed to read test results: ${e}`);
        }
      }
      // Save error log location for user reference
      log(`For detailed logs, check: ${logFile}`);
      process.exit(code);
    }
  });

  // Keep the process alive so the window doesn't close immediately
  process.stdin.resume();
} catch (err) {
  log(`Test execution error: ${err.message}`);
  if (err.stack) {
    log(`Stack trace: ${err.stack}`);
  }
  process.stdin.resume();
  process.exit(1);
}
