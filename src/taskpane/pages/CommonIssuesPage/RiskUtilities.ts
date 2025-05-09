import { Issue } from "./Types";
import { mergeStyles } from "@fluentui/react";

export function getRisk(item: Issue): string {
  let impact = getRiskImpactAndLikelihood(item)?.impact;
  let likelihood = getRiskImpactAndLikelihood(item)?.likelihood;
  const riskMap: { [key: string]: { [key: string]: string } } = {
    High: { High: "C", Medium: "H", Low: "M" },
    Medium: { High: "H", Medium: "M", Low: "L" },
    Low: { High: "M", Medium: "L" },
  };
  const risk = riskMap[String(impact)]?.[String(likelihood)];
  if (risk) {
    return risk;
  }
  return "NA";
}

export function getRiskStyle(item: Issue) {
  class RiskColors {
    backgroundColor: string;
    textColor: string;

    constructor(backgroundcolor: string, textColor: string) {
      this.backgroundColor = backgroundcolor;
      this.textColor = textColor;
    }
  }
  function getRiskColor(item: Issue) {
    let risk = getRisk(item);
    switch (risk) {
      case "C":
        return new RiskColors("Maroon", "White");
      case "H":
        return new RiskColors("Red", "Black");
      case "M":
        return new RiskColors("Orange", "Black");
      case "L":
        return new RiskColors("Yellow", "Black");
      default:
        return new RiskColors("DeepSkyBlue", "Black");
    }
  }
  return mergeStyles({
    width: "16px",
    height: "16px",
    backgroundColor: `${getRiskColor(item).backgroundColor}`,
    textAlign: "center",
    lineHeight: "16px",
    marginRight: "8px",
    color: `${getRiskColor(item).textColor}`,
    fontWeight: "bold",
    fontSize: "10px",
  });
}

export function getRiskImpactAndLikelihood(item: Issue) {
  let fullMarkdown = item.markdown;
  let impactMatch = fullMarkdown.match(/#### Impact:\s*(\w+)/i);
  let likelihoodMatch = fullMarkdown.match(/#### Likelihood:\s*(\w+)/i);
  let impact = null;
  let likelihood = null;
  if (impactMatch && impactMatch.length > 1 && likelihoodMatch && likelihoodMatch.length > 1) {
    impact = impactMatch[1];
    likelihood = likelihoodMatch[1];
    return { impact: impact, likelihood: likelihood };
  }
  return { impact: 0, likelihood: 0 };
}

export function getIssueTooltip(item: Issue) {
  let impact = getRiskImpactAndLikelihood(item)?.impact;
  let likelihood = getRiskImpactAndLikelihood(item)?.likelihood;
  let tooltipText = "Category: " + item.group.replace(/[^a-zA-Z]/g, "") + "\n";
  tooltipText += "Impact: " + impact + "\n";
  tooltipText += "Likelihood: " + likelihood + "\n";
  return tooltipText;
}
