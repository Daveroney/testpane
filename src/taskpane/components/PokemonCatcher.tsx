import { Button } from "@fluentui/react-components";
import React from "react";
import "./pokemonCatcher.css";
import { PokemonApi } from "../../api/PokemonApi";

interface PokemonCatcherProps {
  pokemonApi: PokemonApi;
}

export function PokemonCatcher({ pokemonApi }: PokemonCatcherProps) {
  const [loadingMessage, setLoadingMessage] = React.useState("");
  const [error, setError] = React.useState("");

  let dialog: Office.Dialog;

  async function handleDialog(): Promise<void> {
    try {
      const dialogUrl = `https://localhost:3000/dialog.html`;

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 50, width: 70 },
        async (pokemonResult) => {
          dialog = pokemonResult.value;
          // Will trigger when the child (PokemonDialog) delivers a message to this component.
          // This is done because the dialog exists in another context.
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            async (arg: Office.DialogParentMessageReceivedEventArgs) => {
              const message = JSON.parse(arg.message);
              // We send the available/filtered partners when the dialog is ready to receive data.
              if (message.type === "ready") {
                await sendRandomPokemon();
              }
              if (message.type === "pokemonSelected") {
              }
            }
          );
          dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
        }
      );
    } catch {
      setError("Partnerdaten konnten nicht geladen werden.");
    }
  }

  // Handle errors and events around the dialog, like when the dialog is closed
  function processDialogEvent(arg) {
    // 12006 = Dialog is closed via "X" button
    if (arg.error == 12006) {
      setLoadingMessage("");
    }
  }

  async function sendRandomPokemon() {
    const randomPokemonIndex = Math.floor(Math.random() * (1026 - 1 + 1)) + 1;
    const pokemon = await pokemonApi.getPokemon(randomPokemonIndex);
    dialog.messageChild(JSON.stringify({ type: "pokemonData", data: pokemon }));
  }

  return <Button onClick={handleDialog}>Find a Pokémon</Button>;
}
