function getHyphenEluMap() {
  const dictio = new Map();

  dictio.set('bem-vinda', ['bem-vinde']);
  dictio.set('bem-vindas', ['bem-vindes']);
  dictio.set('bem-vindo', ['bem-vinde']);
  dictio.set('bem-vindos', ['bem-vindes']);
  dictio.set('não-binária', ['não-binárie']);
  dictio.set('não-binárias', ['não-bináries']);
  dictio.set('não-binário', ['não-binárie']);
  dictio.set('não-binários', ['não-bináries']);

  return dictio;
}
