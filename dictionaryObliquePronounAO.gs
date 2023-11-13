function getObliquePronounAOMap() {
  const dictio = new Map();

  dictio.set('o', ['a/o', 'o/a']);
  dictio.set('os', ['as/os', 'os/as']);
  dictio.set('lo', ['la/o', 'lo/a']);
  dictio.set('los', ['las/os', 'los/as']);
  dictio.set('no', ['na/o', 'no/a']);
  dictio.set('nos', ['nas/os', 'nos/as']);

  return dictio;
}
