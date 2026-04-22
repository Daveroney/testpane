export class PokemonApi {
  private _apiUrl: string;

  constructor(apiUrl: string) {
    this._apiUrl = apiUrl;
  }

  public async getPokemon(pokemonIndex: number): Promise<any> {
    try {
      const response = await fetch(`${this._apiUrl}/${pokemonIndex}`);
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return await response.json();
    } catch (error) {
      console.error("Error fetching Pokemon data:", error);
      throw error;
    }
  }
}
