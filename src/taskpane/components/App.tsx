import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";
import { PokemonCatcher } from "./PokemonCatcher";
import { PokemonApi } from "../../api/PokemonApi";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const apiUrl = "https://pokeapi.co/api/v2/pokemon/";
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: HeroListItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];
  const pokemonApi = new PokemonApi(apiUrl);

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <HeroList message="Discover which Pokémon you can catch!" items={listItems} />
      <PokemonCatcher pokemonApi={pokemonApi} />
    </div>
  );
};

export default App;
