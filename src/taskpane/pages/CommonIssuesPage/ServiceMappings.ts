interface ServiceMap {
  name: string;
  secureName: string;
  defaultPort: string;
  defaultSecurePort: string;
}

export const ServiceMappings: Map<string, ServiceMap> = new Map([
  [
    "ftp",
    {
      name: "ftp",
      secureName: "ftps",
      defaultPort: "21",
      defaultSecurePort: "990",
    },
  ],
  [
    "ftp?",
    {
      name: "ftp",
      secureName: "ftps",
      defaultPort: "21",
      defaultSecurePort: "990",
    },
  ],
  [
    "ssh",
    {
      name: "ssh",
      secureName: "ssh",
      defaultPort: "22",
      defaultSecurePort: "22",
    },
  ],
  [
    "telnet",
    {
      name: "telnet",
      secureName: "INVALID",
      defaultPort: "23",
      defaultSecurePort: "INVALID",
    },
  ],
  [
    "tftp",
    {
      name: "tftp",
      secureName: "INVALID",
      defaultPort: "69",
      defaultSecurePort: "INVALID",
    },
  ],
  [
    "msrdp",
    {
      name: "rdp",
      secureName: "rdp",
      defaultPort: "3389",
      defaultSecurePort: "3389",
    },
  ],
  [
    "www",
    {
      name: "http",
      secureName: "https",
      defaultPort: "80",
      defaultSecurePort: "443",
    },
  ],
]);
