---
name: coded-workflow-architect
description: |
  Comprehensive workflow for generating and editing Coded Workflows (C# .cs files) in UiPath Studio Desktop.
  Use this when users need to create new C# automations, modify existing coded workflows, fix C# errors,
  or iterate on coded workflow implementations. Supports discovery-first approach with error-driven refinement.
icon: FaCode
color: "#0078D4"
---

# Coded Workflow Architect

Generate and edit Coded Workflows (C# .cs files) using a **discovery-first approach** with **iterative error-driven refinement**. Always understand available APIs before acting, start simple, and validate continuously.

## Core Principles

1. **API Discovery Before Generation** - Never generate C# code without first understanding the available APIs. Search through the project for .cs files and if you find 5+ such files use them as examples first. This ensures that the Coded Workflows throughout the project are consistent. In case there are not sufficient files in the project or the information from the files is not enough, then call CodeGenerationPrerequisitesTool to understand the available APIs.
2. **Start Simple, Iterate** - Create minimal working version first, then refine through validation cycles
3. **Validate After Every Change** - Never assume success; always check with GetErrorsTool
4. **Fix Errors Methodically** - Categorize errors and fix in order: Syntax → Type → Logic
5. **Use Inherited Services** - Leverage services inherited from CodedWorkflow base class (excel, mail, system, testing, uiAutomation, workflows, connections)

---

## Tool Quick Reference

| Tool | Purpose | Key Parameters |
|------|---------|----------------|
| **FileSearchTool** | Find .cs files by regex pattern (MANDATORY first step) | `regexQuery`, `rootDirectory`, `maxResults` |
| **ReadFileTool** | Read file contents with line numbers | `filePath`, `offset`, `limit` |
| **WriteFileTool** | Create new file | `filePath`, `content` |
| **EditFileTool** | Edit existing file via string replacement | `filePath`, `edits[]` |
| **GetErrorsTool** | Check for compilation errors | `onlyCurrentFile` |
| **GetQuickFixesTool** | Get quick fix suggestions for errors | `filePath` |
| **GetTypeDefinitionsTool** | Get type info at specific location | `filePath`, `line`, `columnStart`, `columnEnd`, `searchQuery` |
| **GetProjectContextTool** | Get project info including Object Repository | `queryType` (`full`, `objects`, `entities`, etc.) |
| **RunWorkflowTool** | Run/debug a workflow file | `filePath` |
| **CodeGenerationPrerequisitesTool** | Get APIs (ONLY if <5 .cs files OR no relevant examples) | `userRequest` |

---

## Workflow: Classify Request

**Determine CREATE or EDIT before proceeding:**

| Request Type | Trigger Words | Action |
|--------------|---------------|--------|
| **CREATE** | "generate", "create", "make", "build", "new" | Start with Discovery → Generate |
| **EDIT** | "update", "change", "fix", "modify", "add to" | Start with Discovery → Edit |

If unclear which file to edit, **ask the user** rather than guessing.

---

## Phase 1: Discovery

**Goal:** Understand available APIs and project context before writing any code.

⚠️ **CRITICAL REQUIREMENT**: You MUST search for and analyze existing .cs files in the project BEFORE calling CodeGenerationPrerequisitesTool. Step 1.2 is MANDATORY and NOT OPTIONAL.

### Step 1.1: Read Current File (for EDIT) or Check Project Structure

For **EDIT** requests:
```
ReadFileTool:
  filePath: <workflow to edit>
```

For **CREATE** requests:
- Check project structure to determine appropriate file location
- Identify naming conventions from existing files

### Step 1.2: API Discovery

**MANDATORY: Before generating any C# code, learn from existing project patterns first.**

**Step A: Search for Existing C# Files**

First, use FileSearchTool to find all .cs files in the project:

```
FileSearchTool:
  regexQuery: ".*\\.cs$"
  rootDirectory: <absolute path to project root - get from context or ask user>
  maxResults: 100
```

**How to get rootDirectory:**
- Use the project path from context (e.g., from GetProjectContextTool or session context)
- Or ask the user for the absolute path if not available
- Example: `C:\Users\TestUser\Documents\UiPath\MyProject\`

**After you get the results:**
- Count the total number of .cs files returned
- Mentally exclude files in `./local/` and `./codedworkflows/` directories from your count
- These folders contain generated/temporary files that should not count as examples

**Step B: Analyze Existing Files (if at least 5 .cs files found)**

If you found **5 or more .cs files**, READ at least 5 of them (preferably diverse examples):

```
ReadFileTool:
  filePath: <path to first .cs file>

ReadFileTool:
  filePath: <path to second .cs file>

# ... read at least 5 files total
```

**Extract and document:**
- Common using statements (e.g., `using UiPath.CodedWorkflows;`)
- Namespace patterns (e.g., `namespace ProjectName`)
- Class structure (inheritance from `CodedWorkflow`)
- Service usage patterns (e.g., `excel.ReadRange()`, `mail.SendMail()`, `uiAutomation.Click()`)
- Argument patterns (input parameters, return tuples for outputs)
- Logging patterns (e.g., `Log("message")`)
- Error handling patterns (try-catch blocks)

**This ensures your generated code is consistent with the project's existing style.**

---

🛑 **CHECKPOINT: Before proceeding to Step C, verify you completed:**
- [ ] Used FileSearchTool to search for .cs files (with `regexQuery: ".*\\.cs$"`)
- [ ] Counted the results (excluding `./local/` and `./codedworkflows/`)
- [ ] If 5+ files found, READ at least 5 of them using ReadFileTool
- [ ] Extracted common patterns from the files you read

**If you haven't done ALL of the above, GO BACK and complete Steps A and B first.**

---

**Step C: Use CodeGenerationPrerequisitesTool ONLY if necessary**

⛔ **DO NOT call this tool if you found 5+ .cs files with relevant examples.**

Call this tool **ONLY IF**:
- Fewer than 5 .cs files were found in the project, OR
- The existing files don't contain examples relevant to the user's request

**What counts as "relevant examples"?**
- User wants Excel operations → Files must show `excel.` service usage
- User wants Mail operations → Files must show `mail.` service usage
- User wants UI Automation → Files must show `uiAutomation.` service usage
- User wants general C# workflow → ANY .cs file with CodedWorkflow class is relevant
- **In most cases, existing files ARE relevant** because they show class structure, namespaces, and base patterns

```
CodeGenerationPrerequisitesTool:
  userRequest: <user's request description>
```

This tool returns:
- Available API methods for the requested functionality
- Required namespaces and using statements
- Example code patterns (with generic descriptor examples)
- Type information for parameters and return values

**IMPORTANT: If the tool returns UI Automation code with Descriptors**

The CodeGenerationPrerequisitesTool may return generic examples, you need to replace the strings with actual values from ObjectRepository.cs:
```csharp
app.Click("someButton"); //How it SHOULD NOT look, this is what CodeGenerationPrerequisitesTool returns
app.Click(Descriptors.SomeApp.SomeScreen.SomeButton); //How it SHOULD look, this is an example of what you should generate
```

**You MUST replace these generic descriptors with actual ones from the project's ObjectRepository.cs:**

1. Read the ObjectRepository.cs file:
   ```
   ReadFileTool:
     filePath: <project_directory>/.local/.codedWorkflows/ObjectRepository.cs
   ```

2. Find the actual descriptor paths in the file (e.g., `Descriptors.MyApp.LoginScreen.LoginButton`)

3. Replace the generic descriptor paths in your generated code with the actual ones from ObjectRepository.cs

4. If NO descriptor mathing the user request exists, follow the guidance in Step 1.4 (ask user to create it in Object Repository)

**Decision Tree:**
```
Found >= 5 .cs files?
  ├─ YES: Do existing files show relevant API usage?
  │   ├─ YES: Use patterns from existing files → SKIP CodeGenerationPrerequisitesTool
  │   └─ NO: Call CodeGenerationPrerequisitesTool for specific API guidance
  └─ NO: Call CodeGenerationPrerequisitesTool
```

### Step 1.3: Understand Available Services

Coded workflows inherit from `CodedWorkflow` base class which provides these services:

| Service Property | Type | Purpose |
|------------------|------|---------|
| `excel` | IExcelService | Excel operations (UseExcelFile or UseWorkBook) |
| `mail` | IMailService | Email operations (send, read, search, attachments) |
| `system` | ISystemService | System/Core activities (delays, file operations, etc.) |
| `testing` | ITestingService | Testing and verification operations |
| `uiAutomation` | IUiAutomationAppService | UI Automation (click, type, get element, etc.) |
| `workflows` | WorkflowRunnerService | Invoke other workflows in the project |
| `connections` | ConnectionsManager | Connection management for integrations |

### Step 1.4: Check Object Repository (for UI Automation)

**When the workflow involves UI Automation**, ALWAYS check for existing selectors before generating code. This applies whether you learned from existing .cs files OR used CodeGenerationPrerequisitesTool.

**Option 1: Use GetProjectContextTool (preferred for quick overview):**
```
GetProjectContextTool:
  queryType: "objects"
```

**Option 2: Read ObjectRepository.cs directly (preferred for seeing full descriptor paths):**
```
ReadFileTool:
  filePath: <project_directory>/.local/.codedWorkflows/ObjectRepository.cs
```

Both methods return all UI elements (selectors/descriptors) defined in the project's Object Repository. Use Option 2 when you need to see the exact descriptor paths to replace generic examples from CodeGenerationPrerequisitesTool.

**Decision tree for selectors:**

1. **Matching selector EXISTS** → Use it via `Descriptors.AppName.ScreenName.ElementName`
2. **Similar selector EXISTS** → Consider if it can be reused or if a new one is needed
3. **NO matching selector** → **Ask the user** to create one before proceeding

**When to ask user to create a selector:**

| Scenario | Action |
|----------|--------|
| UI element not in Object Repository | Ask user: "I need a selector for [element description]. Please add it to the Object Repository and let me know the descriptor path." |
| Multiple similar elements exist | Ask user: "I found similar selectors: [list]. Which one should I use, or should I wait for you to create a new one?" |
| Ambiguous element identification | Ask user: "Please clarify which UI element you want to interact with and ensure it's in the Object Repository." |

**Why use Object Repository instead of hardcoded selectors:**
- Centralized selector management
- Easier maintenance when UI changes
- Reusable across multiple workflows
- Better reliability with fuzzy matching
- Supports multiple selector strategies

---

## Phase 2: Generate or Edit

### For CREATE Requests

**Strategy:** Generate minimal working version using discovered APIs.

```
WriteFileTool:
  filePath: <appropriate path based on project structure>
  content: <C# code following proper structure>
```

**File structure requirements:**
- Place in appropriate folder (e.g., `Workflows/`, project root)
- Use `.cs` extension
- Follow project naming conventions
- Include required using statements
- Inherit from `CodedWorkflow`
- Add `[Workflow]` attribute to entry method
- **If using UI Automation**: Replace any generic `Descriptors.*` paths from CodeGenerationPrerequisitesTool with actual paths from ObjectRepository.cs

### For EDIT Requests

**Strategy:** Always read current content before editing.

```
ReadFileTool:
  filePath: <workflow to edit>
```

Then apply changes:

```
EditFileTool:
  filePath: <workflow>
  edits: [
    { oldContent: <exact text from file>, newContent: <modified text> }
  ]
```

**Critical:** `oldContent` must match exactly and be unique in the file. Include surrounding context if needed for uniqueness.

---

## Phase 3: Validate & Fix Loop

**This phase repeats until 0 errors.**

### Step 3.1: Check for Errors

```
GetErrorsTool:
  onlyCurrentFile: true
```

### Step 3.2: Get Quick Fixes (if errors exist)

```
GetQuickFixesTool:
  filePath: <workflow file>
```

Review suggested fixes - many common issues have automated solutions.

### Step 3.3: Get Type Definitions (for type errors)

```
GetTypeDefinitionsTool:
  filePath: <workflow>
  line: <error line>
  columnStart: <start column>
  columnEnd: <end column>
  searchQuery: <type or member name>
```

Use this to understand:
- What properties/methods are available on a type
- Correct parameter types for methods
- Return type information

### Step 3.4: Apply Fixes

```
EditFileTool:
  filePath: <workflow>
  edits: [{ oldContent: <error code>, newContent: <fixed code> }]
```

### Step 3.5: Iteration Loop

```
REPEAT:
  1. GetErrorsTool
  2. IF 0 errors → EXIT to Phase 4
  3. GetQuickFixesTool for suggestions
  4. GetTypeDefinitionsTool for type issues
  5. ReadFileTool to undestend the current content of the file
  6. Apply appropriate fix with EditFileTool
  7. GOTO 1

DO NOT stop until GetErrorsTool returns 0 errors.
DO NOT skip validation steps.
DO NOT assume edits worked without checking.
```

Expect 3-7 iteration cycles for complex workflows.

---

## Phase 4: Run & Test (Optional)

**Goal:** Execute the coded workflow to verify it works correctly.

### Step 4.1: Run the Workflow

After reaching 0 compilation errors, optionally run the workflow to test its execution:

```
RunWorkflowTool:
  filePath: <full path to the .cs workflow file>
```

**When to run:**
- User explicitly requests to test/run the workflow
- Workflow has logic that should be verified at runtime
- User wants to see the output or behavior

**When NOT to run:**
- Workflow requires external dependencies not available (databases, APIs, etc.)
- Workflow performs destructive operations (deleting files, sending emails to real recipients)
- User only asked for code generation, not execution
- Workflow requires user input or UI interaction

### Step 4.2: Handle Runtime Errors

If the workflow fails during execution:

1. **Analyze the error output** - Identify the root cause from the error message
2. **Determine fix location** - Use the stack trace to find the problematic code
3. **Apply fix** - Use EditFileTool to correct the issue
4. **Re-validate** - Run GetErrorsTool to ensure no new compilation errors
5. **Re-run** - Execute RunWorkflowTool again to verify the fix

**Common runtime issues:**
| Issue | Cause | Fix |
|-------|-------|-----|
| NullReferenceException | Accessing null object | Add null checks or initialize objects |
| FileNotFoundException | File path doesn't exist | Verify path or add existence check |
| ArgumentException | Invalid parameter value | Validate inputs before use |
| TimeoutException | Operation took too long | Add timeout handling or increase timeout |
| UnauthorizedAccessException | Permission denied | Check file/resource permissions |

---

## Phase 5: Response

**Provide comprehensive summary:**

1. **File path** of created/edited workflow (clickable reference)
2. **Brief description** of what the workflow does
3. **Key functionality** implemented
4. **APIs used** and their purposes
5. **Execution results** (if workflow was run) - include output or any runtime observations
6. **Limitations** or notes for the user
7. **Suggested next steps** (further testing, adding error handling, parameterization, etc.)

**Do NOT just say "workflow created"** - give user confidence the request was fully fulfilled.

---

## Coded Workflow Reference

### File Format and Structure

```csharp
using System;
using UiPath.CodedWorkflows;
// Additional using statements as needed

namespace ProjectName
{
    public class MyWorkflow : CodedWorkflow
    {
        [Workflow]
        public void Execute()
        {
            // Workflow implementation
        }
    }
}
```

### Available APIs

Access these through inherited service properties:

| API | Property | Common Operations |
|-----|----------|-------------------|
| **Excel** | `excel` | `UseExcelFile`, `UseWorkBook` |
| **Mail** | `mail` | `SendMail`, `ReadMail`, `GetMailFolders`, `SaveAttachments` |
| **System** | `system` | `Delay`, `WriteTextFile`, `ReadTextFile`, `GetEnvironmentVariable` |
| **Testing** | `testing` | `VerifyExpression`, `VerifyAreEqual`, `LogMessage` |
| **UI Automation** | `uiAutomation` | `Click`, `TypeInto`, `GetText`, `ElementExists`, `Open` |

### Object Repository Usage

Access UI elements defined in the Object Repository via the `Descriptors` static class:

```csharp
// Pattern: Descriptors.AppName.ScreenName.ElementName
var emailField = Descriptors.UiPath_Banking_App.Form.Email;
var submitButton = Descriptors.MyApp.LoginScreen.SubmitButton;

// Use with UI Automation
uiAutomation.TypeInto(emailField, "user@example.com");
uiAutomation.Click(submitButton);
```

### Workflow Invocation

Invoke other workflows in the project via the `workflows` property:

```csharp
// Invoke a workflow with no arguments
workflows.ProcessData();

// Invoke a workflow with arguments
var result = workflows.CalculateTotal(100, 0.1m);

// Invoke and capture multiple return values
var (success, message) = workflows.ValidateInput(inputData);
```

### Arguments (In/Out/InOut Patterns)

**Input Arguments:** Method parameters

```csharp
[Workflow]
public void Execute(string inputPath, int maxRetries)
{
    // inputPath and maxRetries are input arguments
}
```

**Output Arguments:** Return values (use tuples for multiple outputs)

```csharp
[Workflow]
public (bool success, string message) Execute(string input)
{
    // Process input
    return (true, "Completed successfully");
}
```

**InOut Arguments:** Specific use for Coded Workflows, syntax deffers for one single InOut arguments or multiple

```csharp
// Single in argument named Output becames an InOut argument
[Workflow]
public int Execute(int Output) 
{
    Output++;
    return Output;
}

// In case of multiple InOut arguments, then a the retun type needs to be a tuple
//and the names and types of the argumnets need to match 
[Workflow]
public (int count, bool isDone) Execute(int count, bool isDone)
{
    count++;
    isDone = false;
    return (count, isDone);
}
```

### Logging Best Practices

Use the inherited `Log` method for consistent logging:

```csharp
Log("Starting workflow execution...");
Log($"Processing {items.Count} items");
Log("Workflow completed successfully", LogLevel.Info);
Log("Warning: Retry limit approaching", LogLevel.Warn);
Log("Error occurred: " + ex.Message, LogLevel.Error);
```

---

## Code Examples

### Basic Workflow Template

```csharp
using System;
using UiPath.CodedWorkflows;

namespace MyProject
{
    public class BasicWorkflow : CodedWorkflow
    {
        [Workflow]
        public void Execute()
        {
            Log("Starting workflow...");

            // Your implementation here

            Log("Workflow completed.");
        }
    }
}
```

### Workflow with Arguments

```csharp
using System;
using UiPath.CodedWorkflows;

namespace MyProject
{
    public class ProcessDataWorkflow : CodedWorkflow
    {
        [Workflow]
        public (int processedCount, string status) Execute(string inputFile, bool validateData)
        {
            Log($"Processing file: {inputFile}");

            int count = 0;
            string status = "Success";

            try
            {
                // Process the file
                if (validateData)
                {
                    // Validation logic
                }
                count = 10; // Example processed count
            }
            catch (Exception ex)
            {
                status = $"Error: {ex.Message}";
                Log(ex.Message, LogLevel.Error);
            }

            return (count, status);
        }
    }
}
```

### Using Excel API

```csharp
using System;
using System.Data;
using UiPath.CodedWorkflows;

namespace MyProject
{
    public class ExcelWorkflow : CodedWorkflow
    {
        [Workflow]
        public void Execute(string excelPath)
        {
            Log("Reading Excel file...");

            // Read data from Excel
            // Specify the path to your Excel file
            string excelFilePath = "Data.xlsx"; // ask the user to pride the actual path if you are unsure
            
            // Open the Excel file and read data from it
            using (var workbook = excel.UseExcelFile(excelFilePath))
            {
                // Read data from Sheet1 into a DataTable
                // Parameters: hasHeaders (true = first row is header), visibleRowsOnly (true = skip hidden rows)
                DataTable dataTable = workbook.Sheet["Sheet1"].ReadRange(true, true);
                
                // Log the number of rows read
                Log($"Successfully read {dataTable.Rows.Count} rows from the Excel file.");
                
                // Display each row from the DataTable
                foreach (DataRow row in dataTable.Rows)
                {
                    string rowContent = string.Join(" | ", row.ItemArray);
                    Log(rowContent);
                }
            }

            Log("Excel processing completed.");
        }
    }
}
```

### Using Mail API

```csharp
using System;
using System.Collections.Generic;
using UiPath.CodedWorkflows;

namespace MyProject
{
    public class MailWorkflow : CodedWorkflow
    {
        [Workflow]
        public void Execute(string recipient, string subject, string body)
        {
            var mailOptions = new SendOutlookMailOptions
            {
                // Set the recipient's email address
                To = new List<string> { recipient },

                // Set the subject of the email
                Subject = subject,

                // Set the body of the email
                Body = body
            };
            mail.Outlook().SendMail(mailOptions);
            
            Log($"Email sent to {recipient}");
        }
    }
}
```

### Using UI Automation API

```csharp
using System;
using UiPath.CodedWorkflows;

namespace MyProject
{
    public class UIAutomationWorkflow : CodedWorkflow
    {
        [Workflow]
        public void Execute(string username, string password)
        {
            Log("Starting UI automation...");

            // Open application
            var app = uiAutomation.Open(Descriptors.MyApp.Application);

            // Type into fields using Object Repository descriptors
            app.TypeInto(Descriptors.MyApp.LoginScreen.Username, username);
            app.TypeInto(Descriptors.MyApp.LoginScreen.Password, password);

            // Click login button
            app.Click(Descriptors.MyApp.LoginScreen.LoginButton);

            // Wait for and verify success
            var welcomeText = app.GetText(Descriptors.MyApp.Dashboard.WelcomeMessage);
            Log($"Login successful: {welcomeText}");
        }
    }
}
```

### Invoking Other Workflows

```csharp
using System;
using UiPath.CodedWorkflows;

namespace MyProject
{
    public class OrchestratorWorkflow : CodedWorkflow
    {
        [Workflow]
        public void Execute(string[] items)
        {
            Log($"Processing {items.Length} items...");

            foreach (var item in items)
            {
                // Invoke another workflow for each item
                var (success, result) = workflows.ProcessSingleItem(item);

                if (!success)
                {
                    Log($"Failed to process: {item}", LogLevel.Warn);
                    continue;
                }

                Log($"Processed {item}: {result}");
            }

            // Final cleanup workflow
            workflows.Cleanup();

            Log("All items processed.");
        }
    }
}
```

---

## Common Issues & Fixes

| Issue | Cause | Fix |
|-------|-------|-----|
| `CS0246: Type not found` | Missing using statement | Add appropriate `using` directive |
| `CS1061: Does not contain definition` | Wrong method name or missing service | Check existing .cs files for correct API patterns, or use GetTypeDefinitionsTool |
| `CS0029: Cannot convert type` | Type mismatch | Use GetTypeDefinitionsTool to check expected types |
| `CS0103: Name does not exist` | Undefined variable/property | Declare variable or check spelling |
| `CS0161: Not all paths return value` | Missing return statement | Add return for all code paths |
| `CS0128: Local variable already defined` | Duplicate variable name | Rename or remove duplicate |
| `[Workflow] not recognized` | Missing UiPath.CodedWorkflows reference | Ensure project has correct packages |
| Service property not available | Not inheriting from CodedWorkflow | Add `: CodedWorkflow` to class declaration |
| Object Repository element not found | Wrong descriptor path | Check exact path in Object Repository |
| `Descriptors` does not contain element | Selector not in Object Repository | Ask user to add the element to Object Repository |
| UI element not found at runtime | Selector doesn't match actual UI | Ask user to update/re-capture the selector |

---

## Anti-Patterns

**Never:**
- Generate C# code without first searching for existing .cs files in the project (Step 1.2)
- Skip analyzing existing .cs files when 5+ are available in the project
- Call CodeGenerationPrerequisitesTool when existing files already show relevant patterns
- Assume a create/edit succeeded without validating with GetErrorsTool
- Stop iteration loop before reaching 0 errors
- Guess API method names - always verify with discovery tools
- Edit without reading current file content first
- Use generic `oldContent` that matches multiple locations
- Skip the `[Workflow]` attribute on the entry method
- Forget to inherit from `CodedWorkflow` base class
- Hardcode UI selectors instead of using Object Repository descriptors
- Generate UI Automation code without first checking Object Repository for existing selectors
- Assume selectors exist without verifying with GetProjectContextTool
- Ignore GetQuickFixesTool suggestions for common errors
- Run workflows with compilation errors (always reach 0 errors first)
- Run workflows that perform destructive operations without user confirmation
- Run workflows that require unavailable external dependencies

---

## Quality Checklist

Before handover, verify:

- [ ] GetErrorsTool returns 0 errors
- [ ] File path is valid and follows project conventions
- [ ] Class inherits from `CodedWorkflow`
- [ ] Entry method has `[Workflow]` attribute
- [ ] All using statements are present
- [ ] Arguments follow proper In/Out patterns
- [ ] Logging is included for key operations
- [ ] Error handling (try-catch) is included where appropriate
- [ ] All UI elements use Object Repository descriptors (no hardcoded selectors)
- [ ] Missing selectors were flagged and user was asked to create them
- [ ] Workflow was run with RunWorkflowTool (if requested or appropriate)
- [ ] Runtime errors were addressed (if workflow was executed)
- [ ] User has been informed of any limitations
- [ ] Next steps have been suggested (testing, customization)
