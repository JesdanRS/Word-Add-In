import { GitLabAuthProvider } from "./GitLabAuthProvider";

interface AppProvidersProps {
  children?: React.ReactNode;
}

export const AppProviders = ({ children }: AppProvidersProps): JSX.Element => {
  return <GitLabAuthProvider>{children}</GitLabAuthProvider>;
};
