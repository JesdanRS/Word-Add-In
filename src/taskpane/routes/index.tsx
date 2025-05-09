import { Route, Routes } from "react-router-dom";
import { AppLayout } from "../layout";
import { CommonIssuesPage } from "../pages/CommonIssuesPage";
import { NotFoundPage } from "../pages/NotFoundPage";

export const AppRoutes = (): JSX.Element => {
  return (
    <Routes>
      <Route path="*" element={<NotFoundPage />} />
      <Route path="" element={<AppLayout />}>
        <Route path="/" element={<CommonIssuesPage />} />
      </Route>
    </Routes>
  );
};
