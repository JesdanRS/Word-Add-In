export interface Issue {
  title: string;
  group: string;
  category: string;
  cvss: string;
  comment?: string;
  markdown: string;
  generation?: {
    getIssueContext?: string;
    getAffectedSystems?: string;
    getPlaceholderContent?: string;
  };
  report_on_external_network_test?: boolean;
  path: string;
}

export enum Risk {
  Informational = 1,
  Low = 2,
  Medium = 3,
  High = 4,
  Critical = 5,
}

export enum Pluralisation {
  Plural = 1,
  Singular = 2,
}
