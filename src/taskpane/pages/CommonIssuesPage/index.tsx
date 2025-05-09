import {
  DefaultButton,
  Dialog,
  DialogFooter,
  GroupedList,
  IconButton,
  PrimaryButton,
  SelectionMode,
  Shimmer,
  ShimmerElementType,
  Spinner,
  SpinnerSize,
  Stack,
  TextField,
  Toggle,
  MessageBar,
  MessageBarType,
  SearchBox,
  IGroup,
  TooltipHost,
  ProgressIndicator,
  Link,
} from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import { useCallback, useContext, useEffect, useState } from "react";
import { GetCommonIssuesLastCommit, GetCommonIssues } from "./GitLabHelper";
import { insertGeneratedIssues } from "./IssueGeneration";
import { Issue } from "./Types";
import { insertIssue } from "./WordHelper";
import * as RiskUtils from "./RiskUtilities";
import { AuthContext } from "react-oidc-context";

export const CommonIssuesPage = (): JSX.Element => {
  const auth = useContext(AuthContext);

  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isBusy, setIsBusy] = useState<boolean>(false);

  const [issues, _setIssues] = useState<Issue[]>(JSON.parse(localStorage.getItem("issues") ?? "[]"));
  const setIssues = useCallback((value: Issue[]) => {
    localStorage.setItem("issues", JSON.stringify(value));
    _setIssues(value);
  }, []);

  const [groups, _setGroups] = useState<IGroup[]>(JSON.parse(localStorage.getItem("groups") ?? "[]"));
  const setGroups = useCallback((value: IGroup[]) => {
    localStorage.setItem("groups", JSON.stringify(value));
    _setGroups(value);
  }, []);

  const [filter, _setFilter] = useState<string>(localStorage.getItem("filter") ?? "");
  const setFilter = useCallback((value: string) => {
    localStorage.setItem("filter", value);
    _setFilter(value);
  }, []);

  const [selectedIssue, setSelectedIssue] = useState<Issue | undefined>(undefined);
  const [affectedSystems, setAffectedSystems] = useState<string>("");

  const [errorText, setErrorText] = useState<string>("");

  const [showInsertDialog, { toggle: toggleInsertDialog }] = useBoolean(true);

  const [showErrorDialog, { toggle: toggleErrorDialog }] = useBoolean(true);

  const TEMPLATE_ISSUE: Issue = {
    title: "Template",
    group: "Template",
    category: "TODO",
    cvss: "{==TODO==}",
    comment: null,
    markdown: `
### {==TODO==}
{==TODO==}
#### Impact: {==TODO==}
{==TODO==}
#### Likelihood: {==TODO==}
{==TODO==}
#### Recommendation
{==TODO==}
`,
    path: "",
  };

  useEffect(() => {
    async function run() {
      let cachedCommit = localStorage.getItem("commit");
      let latestCommit = await GetCommonIssuesLastCommit(auth.user.access_token);

      if (cachedCommit !== latestCommit) {
        let latestIssues = await GetCommonIssues(auth.user.access_token, latestCommit);
        latestIssues.sort((a, b) => {
          const groupComparison = a.group.localeCompare(b.group);
          if (groupComparison !== 0) {
            return groupComparison;
          } else {
            return a.title.localeCompare(b.title);
          }
        });
        let startIndex = 0;
        let latestGroups: IGroup[] = [...new Set(latestIssues.map((x) => x.group))].map((group) => {
          let items = latestIssues.filter((x) => x.group === group);
          startIndex += items.length;
          return {
            count: items.length,
            key: group,
            name: group,
            startIndex: startIndex - items.length,
            isCollapsed: true,
          };
        });
        setIssues(latestIssues);
        setGroups(latestGroups);
        localStorage.setItem("commit", latestCommit);
      }

      setIsLoading(false);
    }

    run();
  }, [auth, setIssues, setGroups]);

  return (
    <>
      <Stack tokens={{ childrenGap: 10 }}>
        {isLoading && (
          <MessageBar messageBarType={MessageBarType.info}>
            Loading common issues from GitLab. This may take a minute...
          </MessageBar>
        )}
        {isLoading && <Spinner size={SpinnerSize.large} />}
        {isBusy && <ProgressIndicator />}
        <DefaultButton
          disabled={isLoading || isBusy}
          text="Insert Template Issue"
          onClick={() => {
            setSelectedIssue(TEMPLATE_ISSUE);
            toggleInsertDialog();
          }}
        />
        <SearchBox
          disabled={isLoading || isBusy}
          placeholder="Filter issue titles using regular expressions"
          value={filter}
          onChange={(_event, newValue) => {
            setFilter(newValue ?? "");
          }}
        />
        {isLoading ? (
          [...Array(10).keys()].map((x) => (
            <Shimmer key={x} shimmerElements={[{ type: ShimmerElementType.line, height: 40 }]} />
          ))
        ) : (
          <GroupedList
            items={issues.filter((x) => {
              try {
                new RegExp(filter ?? "", "i");
                return new RegExp(filter, "i").test(x.title);
              } catch {
                return false;
              }
            })}
            groups={filter !== "" ? undefined : groups}
            selectionMode={SelectionMode.none}
            compact={true}
            onRenderCell={(_nestingDepth?: number, item?: any, _index?: number) => {
              return (
                <Stack horizontal verticalAlign="center">
                  <TooltipHost content="Insert">
                    <IconButton
                      disabled={isBusy}
                      iconProps={{
                        iconName: "insert",
                      }}
                      onClick={() => {
                        setSelectedIssue(item);
                        toggleInsertDialog();
                      }}
                    />
                  </TooltipHost>
                  <TooltipHost
                    content={RiskUtils.getRisk(item)}
                    styles={{
                      root: RiskUtils.getRiskStyle(item),
                    }}
                  >
                    {RiskUtils.getRisk(item)}
                  </TooltipHost>
                  <TooltipHost content={RiskUtils.getIssueTooltip(item)} style={{ whiteSpace: "pre-line" }}>
                    <Link
                      styles={{
                        root: {
                          color: "inherit",
                          textDecoration: "none",
                          cursor: "pointer",
                        },
                      }}
                      onClick={() => {
                        let editUrl = `https://gitlab.com/-/ide/project/mrbl4ck/common-issues/edit/master/-/${item.path}`;
                        Office.context.ui.openBrowserWindow(editUrl);
                      }}
                    >
                      {item.report_on_external_network_test && "*"}
                      {item.title}
                    </Link>
                  </TooltipHost>
                </Stack>
              );
            }}
          ></GroupedList>
        )}
        <MessageBar messageBarType={MessageBarType.info}>
          Web application issues marked with an asterisk should be reported on external network tests.
        </MessageBar>
        <MessageBar messageBarType={MessageBarType.warning}>
          The issue generation feature is for testing and feedback only. Check issues carefully in production reports.
        </MessageBar>
        <DefaultButton
          disabled={isLoading || isBusy}
          text="Insert Generated Issues"
          onClick={() => document.getElementById("InsertGeneratedIssuesFileInput")!.click()}
        />
        <input
          id="InsertGeneratedIssuesFileInput"
          type="file"
          accept=".nessus,.json"
          multiple
          style={{ display: "none" }}
          onClick={(event) => ((event.target as HTMLInputElement).value = "")}
          onChange={async (event) => {
            setIsBusy(true);
            await insertGeneratedIssues(issues, event.target.files);
            setIsBusy(false);
          }}
        />
        <DefaultButton
          disabled={isLoading || isBusy}
          text="Insert All Issues"
          onClick={async () => {
            setIsBusy(true);
            for (
              let i = 0;
              i < (process.env.NODE_ENV === "development" ? issues.slice(0, 10).length : issues.length);
              i++
            ) {
              await insertIssue(issues[i], "{==HOST1==}");
              await insertIssue(issues[i], "{==HOST1==}, {==HOST2==}");
            }
            setIsBusy(false);
          }}
        />
        <DefaultButton
          disabled={isBusy}
          text="Clear Cache"
          onClick={async () => {
            setIsLoading(true);
            localStorage.clear();
            auth.signoutSilent();
          }}
        />
      </Stack>

      <Dialog
        hidden={showErrorDialog}
        onDismiss={toggleErrorDialog}
        dialogContentProps={{
          title: "Error",
          subText: errorText,
        }}
        modalProps={{
          isBlocking: true,
        }}
      ></Dialog>

      <Dialog
        hidden={showInsertDialog}
        onDismiss={toggleInsertDialog}
        dialogContentProps={{
          title: "Insert Issue",
          subText: "Enter affected systems using line, comma, or space separators:",
        }}
        modalProps={{
          isBlocking: true,
        }}
      >
        <TextField
          value={affectedSystems}
          onChange={(_event, newValue) => {
            setAffectedSystems(newValue ?? "");
          }}
          multiline={true}
        />
        <DialogFooter>
          <PrimaryButton
            text="Insert"
            onClick={async () => {
              setIsBusy(true);
              if (selectedIssue) {
                try {
                  await insertIssue(selectedIssue, affectedSystems);
                } catch (error) {
                  setErrorText(
                    'An error occured when attempting to insert the issue. Please see the console for error details."'
                  );
                  toggleErrorDialog();
                  console.error(JSON.stringify(error as OfficeExtension.Error));
                }
              }
              toggleInsertDialog();
              setIsBusy(false);
            }}
          />
        </DialogFooter>
      </Dialog>
    </>
  );
};
