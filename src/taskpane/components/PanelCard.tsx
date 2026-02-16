import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import { theme } from "../styles/themeTokens";

const useStyles = makeStyles({
  card: {
    border: `1px solid ${theme.border}`,
    borderRadius: theme.radius,
    padding: "14px",
    backgroundColor: theme.surface,
    color: theme.text,
  },
});

export default function PanelCard(props: { children: React.ReactNode }) {
  const styles = useStyles();
  return <div className={styles.card}>{props.children}</div>;
}
