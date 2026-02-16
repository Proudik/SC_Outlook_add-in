import * as React from "react";
import { makeStyles, tokens } from "@fluentui/react-components";

export interface HeaderProps {
  title: string;
  logo: string;
  message?: string; 
  size?: "small" | "default";
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    textAlign: "center",
    gap: "8px",
    width: "100%",
  },

  logo: {
    width: "100%",        // 👈 flex with container
    maxWidth: "420px",    // 👈 upper bound for default
    height: "auto",
    objectFit: "contain",
  },

  logoSmall: {
    maxWidth: "260px",    // 👈 tighter for auth screen
  },

  message: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    margin: 0,
  },
});


const Header: React.FC<HeaderProps> = ({ logo, title, message, size = "default" }) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <img
        src={logo}
        alt={title}
        className={`${styles.logo} ${size === "small" ? styles.logoSmall : ""}`}
      />

      {message ? <h1 className={styles.message}>{message}</h1> : null}
    </div>
  );
};

export default Header;
