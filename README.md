# Word Add-in

## Development

- Install VS Code (`choco install vscode`)
- Install Node.js (`choco install nodejs`)
- Start debugging from VS Code using the "Word Desktop (Edge Chromium)" launch configuration.

To resolve the "Allow localhost loopback for Microsoft Edge WebView" error, run the following as an administrator.

```
npx office-addin-dev-settings appcontainer EdgeWebView --loopback
```

## Common Issue Generation

Common issues may define JavaScript code in the issue frontmatter which is used to automatically generate the issue from tool output.

The functions which can be defined are as follows:

```typescript
type IssueContext = any;

/*
  An optional method which returns an object to be passed to 
  getAffectedSystems and getPlaceholderContent in order to reduce 
  code duplication.
*/
function getIssueContext(): IssueContext;

/*
  A required method which returns a string array of the hosts affected
  by the issue.
*/
function getAffectedSystems(issueContext: IssueContext): string[];

/*
  An option function which returns a string used to complete a placeholder
  denoted by "{==PLACEHOLDER==}" in common issue markdown.
*/
function getPlaceholderContent(issueContext: IssueContext): string;
```

Only the function body is required in the common issue frontmatter, as shown in the following example:

```markdown
---
generation:
  getIssueContext: |
    return [
      { hostName: "host1.example.com", pocString: "curl http://host1.example.com" },
      { hostName: "host2.example.com", pocString: "curl http://host2.example.com" },
    ];
  getAffectedSystems: |
    return issueContext.map(x => x.hostName);
  getPlaceholderContent: |
    return issueContext[0].pocString;
---
```

The functions can make used of the tool output and helper functions contained within the global "issueGeneration" object, as defined in https://gitlab.com/mrbl4ck/word-add-in/-/blob/main/src/CommonIssues/IssueGeneration.ts. As this is a global object, the developer tools console can be used for testing generation code.
