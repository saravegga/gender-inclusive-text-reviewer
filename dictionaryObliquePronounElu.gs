function getObliquePronounEluMap() {
  const dictio = new Map();

  dictio.set('a', ['e']);
  dictio.set('as', ['es']);
  dictio.set('o', ['e']);
  dictio.set('os', ['es']);
  dictio.set('la', ['le']);
  dictio.set('las', ['les']);
  dictio.set('lo', ['le']);
  dictio.set('los', ['les']);
  dictio.set('na', ['ne']);
  dictio.set('nas', ['nes']);
  dictio.set('no', ['ne']);
  dictio.set('nos', ['nes']);

  return dictio;
}
