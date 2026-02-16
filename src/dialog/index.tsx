import * as React from "react";
import { createRoot } from "react-dom/client";
import DialogAuth from "./DialogAuth";
import GraphAuthDialog from "./GraphAuthDialog";

declare const Office: any;

const el = document.getElementById("root");
if (!el) throw new Error("Missing root element");

const root = createRoot(el);

function isGraphMode(): boolean {
  try {
    const url = new URL(window.location.href);
    const mode = String(url.searchParams.get("mode") || "").toLowerCase();
    if (mode === "graph") return true;

    const h = String(window.location.hash || "").toLowerCase();
    return h.includes("graph");
  } catch {
    return false;
  }
}

function App() {
  return isGraphMode() ? <GraphAuthDialog /> : <DialogAuth />;
}

function render() {
  root.render(<App />);
}

try {
  if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
    Office.onReady(() => render());
  } else {
    render();
  }
} catch {
  render();
}