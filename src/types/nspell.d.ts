declare module "nspell" {
  interface NSpellInstance {
    correct(word: string): boolean;
    suggest(word: string): string[];
    add(word: string): NSpellInstance;
  }
  /** Hunspell affix + dictionary text, as shipped by the `dictionary-*` packages. */
  function NSpell(aff: string, dic: string): NSpellInstance;
  export default NSpell;
}
