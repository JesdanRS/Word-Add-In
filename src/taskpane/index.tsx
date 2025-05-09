import { createRoot } from "react-dom/client";
import { HashRouter } from "react-router-dom";
import { AppProviders } from "./providers";
import { AppRoutes } from "./routes";
import { initializeIcons } from "@fluentui/react";

/* global document, Office, module, require */

// https://github.com/jonschlinkert/gray-matter/pull/132
global.Buffer = global.Buffer || require("buffer").Buffer;

const rootElement: HTMLElement = document.getElementById("container");
const root = createRoot(rootElement);

/* Render application after Office initializes */
Office.onReady(() => {
  // Initialise Fluent UI icons
  initializeIcons();

  OfficeExtension.config.extendedErrorLogging = process.env.NODE_ENV === 'development';

  root.render(
    <AppProviders>
      <HashRouter>
        <AppRoutes />
      </HashRouter>
    </AppProviders>
  );
});

// if ((module as any).hot) {
//   (module as any).hot.accept("./components/App", () => {
//     const NextApp = require("./components/App").default;
//     root.render(NextApp);
//   });
// }
