import MarkdownIt from "markdown-it";
import { Issue, Risk, Pluralisation } from "./Types";
import { CustomStylesJson } from "./Styles";

function parseAffectedSystemsToHtmlString(affectedSystems: string): string {
  // Split the provided string on commas and whitespace
  let hosts = affectedSystems.split(/,\s*|[\t\n\r\f\v]+/);
  // Remove any empty values
  hosts = hosts.filter((x) => x !== "");
  // Deduplicate
  hosts = [...new Set(hosts)];
  // Sort
  hosts.sort();
  // Loop over hosts and try and parse URIs, turning any valid URIs into HTML anchors
  for (let i = 0; i < hosts.length; i++) {
    try {
      let uri = new URL(hosts[i]);
      if (uri.protocol !== "" && uri.host !== "") {
        hosts[i] = `<a href='${hosts[i]}'>${hosts[i]}</a>`;
      }
    } catch {}
  }
  // If hosts is empty, create a TODO entry
  if (hosts.length === 0) {
    hosts[0] = "{==TODO==}";
  }

  // Return the parsed HTML string
  return hosts.join(", ");
}

export async function insertIssue(issue: Issue, affectedSystems: string) {
  // Parse affected systems
  let affectedSystemsHtml = parseAffectedSystemsToHtmlString(affectedSystems);
  // Use affected systems to determine pluralisation
  let pluralisation = affectedSystemsHtml.indexOf(",") > -1 ? Pluralisation.Plural : Pluralisation.Singular;
  // Use a regex to parse issue pluralisation
  let markdown = issue.markdown.replace(/\{([^{}]*?)\|\|([^{}]*?)\}/g, `$${pluralisation}`);

  let impactRiskString = (markdown.match(/#### Impact: (.*)/) ?? ["Low"])[1];
  let impactRisk = (Risk as any)[impactRiskString];
  let likelihoodRiskString = (markdown.match(/#### Likelihood: (.*)/) ?? ["Low"])[1];
  let likelihoodRisk = (Risk as any)[likelihoodRiskString];
  let risk = impactRisk + likelihoodRisk - 3;
  let riskString = Object.keys(Risk).find((x) => (Risk as any)[x] === risk);
  riskString = riskString === undefined ? "{==TODO==}" : riskString;
  let categoryString = issue.category === "TODO" ? "{==TODO==}" : issue.category;

  // TODO: Change link style, can be removed when common issues updated
  markdown = markdown.replace(/\[(.*?)\]\(\)/g, "<$1>");

  let md = new MarkdownIt();
  let html = md.render(markdown);

  var htmlLines = html.split("\n");
  var tableHtml = `<table><thead><tr><th style="width: 38%"><b>Risk: ${riskString}</b></th><th style="width: 15%"><b>CVSS: ${issue.cvss}</b></th><th style="width: 47%"><b>Category: ${categoryString}</b></th></tr></thead><tbody><tr><td colspan="3"><b>Affected Systems:</b> ${affectedSystemsHtml}</td></tr></tbody></table>`;
  htmlLines.splice(1, 0, tableHtml);
  html = htmlLines.join("\n");

  html = `<head><style></style></head><body>${html}<p></p></body>`;
  // Insert br tags to preserve emtpy paragraphs
  html = html.replace(/<\/(p|pre|table)>\n<(p|pre|table)>/g, "</$1>\n<p><br></p>\n<$2>");

  await Word.run(async (context) => {
    // Queue importing custom styles 
    const styles = context.document.importStylesFromJson(CustomStylesJson);

    // Get the current select and replace our HTML
    const originalRange = context.document.getSelection();
    const insertedRange = originalRange.insertHtml(html, Word.InsertLocation.replace);

    // Queue loading the inserted content for further processing
    insertedRange.load(["paragraphs", "tables"])

    // Queue searching for highlights
    const highlightRanges = insertedRange.search("[{]==*==[}]", {
      matchWildcards: true,
    });
    highlightRanges.load("font, text");

    // Queue loading hyperlinks for later processing
    const hyperlinkRanges = insertedRange.getHyperlinkRanges();
    hyperlinkRanges.load("hyperlink");

    // Synchronise the context - this will make visible changes in the document and load our queued items
    await context.sync();

    // Set the info table style, this will always be the first table in the insert
    insertedRange.tables.items[0].style = "Table Grid Light";
    insertedRange.tables.items[0].autoFitWindow();
    insertedRange.tables.items[0].horizontalAlignment = Word.Alignment.left;

    // Set remaining tables to MrBl4ck style
      insertedRange.tables.items.slice(1).forEach((x) => (x.style = "MrBl4ck"));

    // Set Normal (Web) style to Normal
    insertedRange.paragraphs.items.filter((x) => x.style === "Normal (Web)").forEach((x) => (x.style = "Normal"));

    // Set Heading 3 style to Issue Heading
      insertedRange.paragraphs.items
      .filter((x) => x.style === "Heading 3")
      .forEach((x) => {
        x.style = "Normal";
        x.style = "Issue Heading";
      });

    // Set Heading 4 style to Issue SubHeading
      insertedRange.paragraphs.items
      .filter((x) => x.style === "Heading 4")
      .forEach((x) => (x.style = "Issue SubHeading"));

    // Set HTML Preformatted style to Codify
      insertedRange.paragraphs.items.filter((x) => x.style === "HTML Preformatted").forEach((x) => (x.style = "Codify"));

    // Replace '{==.*==}' with highlights
    highlightRanges.items.forEach((x) => {
      x.font.highlightColor = "yellow";
      x.insertText(x.text.replace("{==", "").replace("==}", ""), Word.InsertLocation.replace);
    });

    // Reset hyperlink content to fix default formatting which gets broken when reapplying Normal
    hyperlinkRanges.items.forEach((linkRange) => {
      linkRange.hyperlink = linkRange.hyperlink;
    });

    // Insert comment
    if (issue.comment) {
      insertedRange.paragraphs.items[0].getRange().insertComment(issue.comment)
    }

    // Move cursor to end of inserted content
    insertedRange.paragraphs.items[insertedRange.paragraphs.items.length - 1].getNext().select(Word.SelectionMode.start);

    // Return and synchronise context to update document
    return await context.sync();
  });
}
