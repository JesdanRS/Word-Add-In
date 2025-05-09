import { Button, Spinner } from "@fluentui/react";
import { User } from "oidc-client-ts";
import { useEffect, useState } from "react";
import { AuthProvider, hasAuthParams, useAuth } from "react-oidc-context";

interface WrapperProps {
  children?: React.ReactNode;
}

const Wrapper = ({ children }: WrapperProps): JSX.Element => {
  const auth = useAuth();
  const [hasTriedSignin, setHasTriedSignin] = useState(false);

  useEffect(() => {
    if (!hasAuthParams() && !auth.isAuthenticated && !auth.activeNavigator && !auth.isLoading && !hasTriedSignin) {
      console.log("Iniciando redirección para autenticación...");
      auth.signinRedirect().catch((error) => {
        console.error("Error en la redirección de autenticación:", error);
      });
      setHasTriedSignin(true);
    }
  }, [auth, hasTriedSignin]);

  switch (auth.activeNavigator) {
    case "signinSilent":
      return <div>Signing you in...</div>;
    case "signoutRedirect":
      return <div>Signing you out...</div>;
  }

  if (auth.isLoading) {
    return <Spinner />;
  }

  if (auth.error) {
    console.error("Error de autenticación:", auth.error);
    if (auth.error.message === "Session not active") {
      console.log("Sesión no activa, redirigiendo...");
      auth.signinRedirect().catch(error => {
        console.error("Error al redirigir:", error);
      });
      return <Spinner />;
    }
    return (
      <div style={{ padding: '20px', textAlign: 'center' }}>
        <div style={{ marginBottom: '10px', color: 'red' }}>
          Error de autenticación: {auth.error.message}
        </div>
        <Button onClick={() => auth.signinRedirect()}>Intentar de nuevo</Button>
      </div>
    );
  }

  if (auth.isAuthenticated) {
    return <>{children}</>;
  }

  return <>Unauthorised</>;
};

interface GitLabAuthProviderProps {
  children?: React.ReactNode;
}

export const GitLabAuthProvider = ({ children }: GitLabAuthProviderProps): JSX.Element => {
  return (
    <AuthProvider
      authority={`${process.env.GITLAB_AUTH_PROVIDER_AUTHORITY}`}
      client_id={`${process.env.GITLAB_AUTH_PROVIDER_CLIENT_ID}`}
      redirect_uri={`${window.location.protocol}//${window.location.hostname}${
        window.location.port ? ":" + window.location.port : ""
      }/word-add-in`}
      onSigninCallback={(_user: User | void): void => {
        console.log("Autenticación completada, procesando callback...");
        // Preservar el hash para mantener la navegación de la aplicación
        window.history.replaceState({}, document.title, `${window.location.pathname}${window.location.hash}`);
        // Forzar actualización del estado de autenticación
        setTimeout(() => {
          window.location.reload();
        }, 500);
      }}
      scope="openid read_api read_repository"
      automaticSilentRenew={true}
      loadUserInfo={true}
    >
      <Wrapper>{children}</Wrapper>
    </AuthProvider>
  );
};
