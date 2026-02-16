import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import SideKickCard from "./SideKickCard";

const useStyles = makeStyles({
  cardWrap: {
    width: "100%",
    maxWidth: "420px",
    padding: "0 12px",
    boxSizing: "border-box",
  },

  stack: {
    display: "flex",
    flexDirection: "column",
    gap: "14px",
  },

  title: {
    fontSize: "20px",
    fontWeight: 700,
    margin: 0,
  },

  subtitle: {
    opacity: 0.75,
    fontSize: "13px",
    lineHeight: 1.5,
  },

  actionsRow: {
    display: "flex",
    gap: "10px",
    marginTop: "8px",
  },

  primaryBtn: {
    height: "36px",
    borderRadius: "10px",
    border: "none",
    padding: "0 18px",
    background: "#2563eb",
    color: "white",
    cursor: "pointer",
    fontWeight: 600,
  },

  secondaryBtn: {
    height: "36px",
    borderRadius: "10px",
    border: "1px solid rgba(0,0,0,0.12)",
    padding: "0 18px",
    background: "transparent",
    cursor: "pointer",
  },
});

type Props = {
  onCancel: () => void;
  onSignedIn: (token: string, email: string) => void;
};

export default function MicrosoftLoginView({ onCancel, onSignedIn }: Props) {
  const styles = useStyles();

  const onContinueMock = () => {
    onSignedIn("mock_ms_token", "test@test.com");
  };

  return (
    <div className={styles.cardWrap}>
      <SideKickCard>
        <div className={styles.stack}>
          <h2 className={styles.title}>Přihlášení přes Microsoft</h2>

          <div className={styles.subtitle}>
            Přihlásíte se pomocí vašeho Microsoft účtu. V další verzi zde proběhne
            skutečné ověření přes Microsoft Entra ID.
          </div>

          <div className={styles.actionsRow}>
            <button className={styles.secondaryBtn} onClick={onCancel}>
              Back
            </button>

            <button className={styles.primaryBtn} onClick={onContinueMock}>
              Continue
            </button>
          </div>
        </div>
      </SideKickCard>
    </div>
  );
}
