import React, { useEffect, useState } from "react";
import { Card, Image, Text } from "@fluentui/react-components";

interface PokemonData {
  name: string;
  sprites: {
    front_default: string;
  };
  types: Array<{
    type: {
      name: string;
    };
  }>;
}

export function PokemonDialog() {
  const [pokemon, setPokemon] = useState<PokemonData | null>(null);

  useEffect(() => {
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      async (arg: Office.DialogParentMessageReceivedEventArgs) => {
        const message = JSON.parse(arg.message);
        if (message.type === "pokemonData") {
          setPokemon(message.data);
        }
      }
    );
    Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
  }, []);

  if (!pokemon) {
    return <div>Loading...</div>;
  }

  return (
    <Card>
      <h1>{pokemon.name}</h1>
      <Image src={pokemon.sprites.front_default} alt={pokemon.name} width={96} height={96} />
      <Text>Types: {pokemon.types.map((t) => t.type.name).join(", ")}</Text>
    </Card>
  );
}
