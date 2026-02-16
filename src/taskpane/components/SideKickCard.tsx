import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import { sidekick } from "../styles/sidekickTokens";

const useStyles = makeStyles({
  card: {
    border: `1px solid ${sidekick.border}`,
    borderRadius: sidekick.radius,
    padding: "14px",
    backgroundColor: sidekick.surface,
    color: sidekick.text,
  },
});

export default function SideKickCard(props: { children: React.ReactNode }) {
  const styles = useStyles();
  return <div className={styles.card}>{props.children}</div>;
}
