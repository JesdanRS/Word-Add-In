import { ServiceMappings } from "./ServiceMappings";
import { Issue } from "./Types";
import { insertIssue } from "./WordHelper";

export interface IssueGenerationWindow extends Window {
  issueGeneration: {
    helpers: {
      groupObjectArrayByKey: (array: any[], key: string) => any;
    };
    nessus: {
      context: Document | null;
      sslPorts: string[];
      helpers: {
        getHostNameFromReportHost: (element: Element) => string;
        getServiceNameFromReportItem: (element: Element) => string;
        getReportItemsForPluginIds: (
          ...ids: string[]
        ) => { hostName: string; serviceName: string; reportItem: Element }[];
      };
    };
    trivy: {
      context: any[] | null;
    };
  };
}

declare let window: IssueGenerationWindow;

window.issueGeneration = {
  helpers: {
    groupObjectArrayByKey: (array: any[], key: string): any => {
      return array.reduce(
        (hash, obj) => ({
          ...hash,
          [obj[key]]: (hash[obj[key]] || []).concat(obj),
        }),
        {}
      );
    },
  },
  nessus: {
    context: null,
    sslPorts: [],
    helpers: {
      getHostNameFromReportHost: (element: Element): string => {
        let name =
          (element.nodeName === "ReportHost" &&
            element.querySelector("HostProperties")!.querySelector("tag[name='host-fqdn']") &&
            element.querySelector("HostProperties")!.querySelector("tag[name='host-fqdn']")!.innerHTML) ||
          (element.nodeName === "ReportHost" &&
            element.querySelector("HostProperties")!.querySelector("tag[name='netbios-name']") &&
            element.querySelector("HostProperties")!.querySelector("tag[name='netbios-name']")!.innerHTML) ||
          (element.nodeName === "ReportHost" && element.getAttribute("name")) ||
          null;

        if (name !== null && name !== undefined) {
          return name!;
        }

        throw new Error("Could not find an appropriate name. Ensure that the provided element is a Nessus ReportHost.");
      },
      getServiceNameFromReportItem: (element: Element): string => {
        if (element.nodeName === "ReportItem") {
          let isSecure = window.issueGeneration.nessus.sslPorts.includes(
            `${(element.parentNode as Element).getAttribute("name")}:${element.getAttribute("port")}`
          );

          let service = element.getAttribute("svc_name");
          let host = window.issueGeneration.nessus.helpers.getHostNameFromReportHost(element.parentNode as Element);
          let port = element.getAttribute("port");

          return `${
            ServiceMappings.has(service!)
              ? isSecure
                ? ServiceMappings.get(service!)!.secureName
                : ServiceMappings.get(service!)!.name
              : service
          }://${host}${
            ServiceMappings.has(service!) &&
            port ===
              (isSecure ? ServiceMappings.get(service!)!.defaultSecurePort : ServiceMappings.get(service!)!.defaultPort)
              ? ""
              : `:${port}`
          }`;
        }

        throw new Error("Could not find an appropriate name. Ensure that the provided element is a Nessus ReportItem.");
      },
      getReportItemsForPluginIds: (
        ...ids: string[]
      ): { hostName: string; serviceName: string; reportItem: Element }[] => {
        return [...window.issueGeneration.nessus.context!.querySelectorAll("ReportItem")]
          .filter((x) => ids.includes(x.getAttribute("pluginID")!))
          .map((x) => {
            return {
              hostName: window.issueGeneration.nessus.helpers.getHostNameFromReportHost(x.parentNode as Element),
              serviceName: window.issueGeneration.nessus.helpers.getServiceNameFromReportItem(x),
              reportItem: x,
            };
          });
      },
    },
  },
  trivy: {
    context: null,
  },
};

export async function insertGeneratedIssues(issues: Issue[], contextFiles: FileList | null) {
  if (contextFiles) {
    await setNessusContext(contextFiles);
    await setTrivyContext(contextFiles);

    for (let i = 0; i < issues.length; i++) {
      try {
        let issue: Issue = JSON.parse(JSON.stringify(issues[i]));

        let issueContext = null;
        if (issue.generation?.getIssueContext !== null && issue.generation?.getIssueContext !== undefined) {
          // eslint-disable-next-line no-new-func
          let getIssueContext = new Function(issue.generation?.getIssueContext!);
          issueContext = getIssueContext();
        }

        if (issue.generation?.getAffectedSystems !== null && issue.generation?.getAffectedSystems !== undefined) {
          // eslint-disable-next-line no-new-func
          let getAffectedSystems = new Function("issueContext", issue.generation?.getAffectedSystems!);

          let affectedSystems = getAffectedSystems(issueContext);

          if (affectedSystems.length > 0) {
            if (
              issue.generation?.getPlaceholderContent !== null &&
              issue.generation?.getPlaceholderContent !== undefined
            ) {
              // eslint-disable-next-line no-new-func
              let getPlaceholderContent = new Function("issueContext", issue.generation?.getPlaceholderContent!);

              issue.markdown = issue.markdown.replace("{==PLACEHOLDER==}", getPlaceholderContent(issueContext));
            }

            await insertIssue(issue, affectedSystems.join(","));
          }
        }
      } catch (e: any) {
        console.error(`Encountered an error processing generation for "${issues[i].title}": ${e.message}`);
      }
    }
  }
}

async function setNessusContext(contextFiles: FileList) {
  let parser = new DOMParser();
  window.issueGeneration.nessus.context = parser.parseFromString("<Files />", "text/xml");

  for (let i = 0; i < contextFiles.length; i++) {
    if (contextFiles[i].name.endsWith(".nessus")) {
      let text = await contextFiles[i].text();
      let xml = parser.parseFromString(text, "text/xml");
      window.issueGeneration.nessus.context.firstChild!.appendChild(xml.firstChild!);
    }
  }

  window.issueGeneration.nessus.sslPorts = [
    ...new Set(
      [...window.issueGeneration.nessus.context.querySelectorAll("ReportItem[pluginID='56984'")].map(
        (x) => `${(x.parentNode as Element).getAttribute("name")}:${x.getAttribute("port")}`
      )
    ),
  ];
}

async function setTrivyContext(contextFiles: FileList) {
  window.issueGeneration.trivy.context = [];

  for (let i = 0; i < contextFiles.length; i++) {
    if (contextFiles[i].name.endsWith(".json")) {
      let text = await contextFiles[i].text();
      try {
        let json = JSON.parse(text);
        if (json.SchemaVersion === 2) {
          window.issueGeneration.trivy.context.push(json);
        }
      } catch {
        console.error(`Could not parse '${contextFiles[i].name}' as json.`);
      }
    }
  }
}
