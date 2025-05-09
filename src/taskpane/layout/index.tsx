import { Stack } from "@fluentui/react";
import { Outlet } from "react-router-dom";

export const AppLayout = (): JSX.Element => {
  return (
    <Stack>
      <Outlet />
    </Stack>
  );
};
