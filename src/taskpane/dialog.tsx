import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { PokemonDialog } from "./components/PokemonDialog/PokemonDialog";

/* global document */

const render = (element: React.ReactNode) => {
  ReactDOM.render(
    <FluentProvider theme={webLightTheme}>{element}</FluentProvider>,
    document.getElementById("dialog-container")
  );
};

Office.onReady(() => {
  render(<PokemonDialog />);
});
