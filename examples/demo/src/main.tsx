import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./demo.css";

const container = document.getElementById("root");
if (!container) throw new Error("demo: #root is missing from index.html");

createRoot(container).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
