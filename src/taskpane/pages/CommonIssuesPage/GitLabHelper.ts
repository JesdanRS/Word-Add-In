import { Gitlab } from "@gitbeaker/rest";
import * as GrayMatter from "gray-matter";
import { Issue } from "./Types";

export async function GetCommonIssuesLastCommit(accessToken: string): Promise<string> {
  const api = new Gitlab({
    host: process.env.GITLAB_URL,
    oauthToken: accessToken,
  });

  let branch = await api.Branches.show(`${process.env.GITLAB_REPO}`, `${process.env.GITLAB_BRANCH}`);

  let commit = branch.commit;

  // https://github.com/microsoft/TypeScript/issues/36981
  return commit.id as string;
}

export async function GetCommonIssues(accessToken: string, commit: string): Promise<Issue[]> {
  const api = new Gitlab({
    host: process.env.GITLAB_URL,
    oauthToken: accessToken,
  });

  let issues = Array<Issue>();

  let tree = await api.Repositories.allRepositoryTrees(`${process.env.GITLAB_REPO}`, {
    path: `${process.env.GITLAB_PATH}`,
    recursive: true,
  });

  tree = tree.filter((x) => /\.md$/.test(x.path));

  await Promise.all(
    tree.map(async (treeItem) => {
      let file = await api.RepositoryFiles.show(`${process.env.GITLAB_REPO}`, treeItem.path, commit);
      try {
        let matter = GrayMatter.default(Buffer.from(file.content, "base64").toString());

        let issue: Issue = {
          title: matter.content.split("\n")[0].replace(/^### (.*)$/, "$1"),
          group: file.file_path
            .split("/")[1]
            .split("_")
            .filter((x) => x.length > 0)
            .map((x) => x.charAt(0).toUpperCase() + x.slice(1))
            .join(" "),
          category: matter.data.category,
          cvss: isNaN(parseFloat(matter.data.cvss))
            ? matter.data.cvss !== "N/A"
              ? `{==${matter.data.cvss}==}`
              : matter.data.cvss
            : parseFloat(matter.data.cvss).toPrecision(matter.data.cvss < 10 ? 2 : 3),
          comment: matter.data.comment,
          markdown: matter.content,
          generation: {
            getIssueContext: matter.data.generation?.getIssueContext,
            getAffectedSystems: matter.data.generation?.getAffectedSystems,
            getPlaceholderContent: matter.data.generation?.getPlaceholderContent,
          },
          report_on_external_network_test: matter.data.report_on_external_network_test,
          path: file.file_path,
        } as Issue;

        issues.push(issue);
      } catch (e: any) {
        console.error(`Encountered an error processing ${file.file_path}`);
        throw e;
      }
    })
  );

  return issues.sort();
}
