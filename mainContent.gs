var leftBoundary = '(^|\\s|\\d|[!"“”#$%&\'()*+,\\-.:;<=>?@[\\]^_`{|}~])';
var rightBoundary = '($|\\s|\\d|[[:punct:]])';

/**
 * @OnlyCurrentDoc
 *
/**
 * Callback for rendering the main card.
 * @return {CardService.Card} The card to show the user.
 */
function onHomepageLoad() {
  try {
    var issueSettingCursor = validateDocAndSetCursorToStart();
    if (issueSettingCursor != null) {
      return issueSettingCursor;
    }

    clearPreviousSuggestions();
    return createStrategyCard();

  } catch(e) {
    Logger.log(e);
    return showGeneralErrorCard();
  }
}

function clearPreviousSuggestions() {
  CacheService.getDocumentCache().remove('changebles');

  var style = {};
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = "#FFFFFF";

  var body = DocumentApp.getActiveDocument().getBody();

  if (body != null && body.getText() != null && body.getText().length > 0) {
    body.editAsText().setAttributes(0, body.getText().length-1, style);
  }
}

function createStrategyCard() {
  try {
    var body = DocumentApp.getActiveDocument().getBody();
    body.editAsText().appendText(" ");
    var cardGrid = CardService.newGrid();

    cardGrid.setBorderStyle(CardService.newBorderStyle().setType(CardService.BorderType.STROKE))
      .setOnClickAction(CardService.newAction().setFunctionName('createSuggestionsCard'))
      .setNumColumns(1)
      .addItem(CardService.newGridItem()
          .setTitle('Terminação "a/o" e "o/a"')
          .setIdentifier('a_o'));

    cardGrid.setBorderStyle(CardService.newBorderStyle().setType(CardService.BorderType.STROKE))
      .setOnClickAction(CardService.newAction().setFunctionName('createSuggestionsCard'))
      .setNumColumns(1)
      .addItem(CardService.newGridItem()
          .setTitle('Terminação "e" + Sistema "elu"')
          .setIdentifier('elu'));
          
    var builder = CardService.newCardBuilder().addSection(buildDescriptionSection());
    builder.addSection(CardService.newCardSection().addWidget(cardGrid));

    return builder.build();

  } catch(e) {
    Logger.log(e);
    return showOpeningDocErrorCard();
  }
}

/**
 * Main function to generate the suggestions card.
 * @return {CardService.Card} The card to show to the user.
 */
function createSuggestionsCard(e) {
  try {
    clearPreviousSuggestions();
    
    var strategy = e.parameters.grid_item_identifier;
    var changebles = getSuggestions(strategy);

    var builder = CardService.newCardBuilder().addSection(buildDescriptionSection());

    builder.addSection(buildTotalSection(changebles.length));
    
    changebles.forEach((changeble) => {
      Logger.log('createSuggestionsCard: changeble word: ' + changeble.word + " / changeble suggestions: " + changeble.suggestions);
      builder.addSection(buildCardSection(changeble));
    });

    return builder.build();

  } catch(e) {
    Logger.log(e);
    return showGeneralErrorCard();
  }
}

function updateCard() {
  var changebles = JSON.parse(CacheService.getDocumentCache().get("changebles"));
  var builder = CardService.newCardBuilder().addSection(buildDescriptionSection());

  builder.addSection(buildTotalSection(changebles.length));

  changebles.forEach((changeble) => {
    builder.addSection(buildCardSection(changeble));
  });

  return CardService.newNavigation().updateCard(builder.build());
}

function buildTotalSection(total) {
  return CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
    .setText("Total de sugestões: " + total + " palavras/expressões"));
}

function buildDescriptionSection() {
  return CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
    .setText("<b>Sugestões de linguagem inclusiva de gênero:</b>"));
}

function buildCardSection(changeble) {
  var cardGrid = CardService.newGrid();  


  for (var i = 0; i < changeble.suggestions.length; i++) {
    
    var suggestion = changeble.suggestions[i];
    if (suggestion == '') continue;

    cardGrid.setBorderStyle(CardService.newBorderStyle().setType(CardService.BorderType.STROKE))
    .setOnClickAction(CardService.newAction().setFunctionName('acceptSuggestion'))
    .setNumColumns(1)
    .addItem(CardService.newGridItem()
        .setTitle(suggestion)
        .setIdentifier(changeble.word + "_" + suggestion + "_"
          + changeble.documentOccurIdx.toString() + "_" + changeble.changebleOccurIdx.toString()));
  }

  return CardService.newCardSection()
            .addWidget(CardService.newTextParagraph().setText('<b><font color="#f58d42">' + changeble.word + '</font></b>'))
            .addWidget(cardGrid)
            .addWidget(CardService.newButtonSet().addButton(CardService.newTextButton()
            .setText('Encontrar')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('findChangebleInDoc')
              .setParameters({word: changeble.word, 
                              documentOccurIdx: changeble.documentOccurIdx.toString()}))
            .setDisabled(false))
            .addButton(CardService.newTextButton()
            .setText('Ignorar')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('ignoreSuggestion')
              .setParameters({word: changeble.word, 
                              documentOccurIdx: changeble.documentOccurIdx.toString(), 
                              changebleOccurIdx: changeble.changebleOccurIdx.toString()}))
            .setDisabled(false)));
}

function getSuggestions(strategy) {
  var changebles = new Array();
  var dictio;
  var dictioDependents;
  var dictioHyphen;
  var dictioObliquePron;
  var dictioNeutral = getNeutralArray();
  
  if (strategy == 'elu') {
    dictio = getEluMap();
    dictioDependents = getDependentsEluMap();
    dictioHyphen = getHyphenEluMap();
    dictioObliquePron = getObliquePronounEluMap();
  } else {
    dictio = getAOMap();
    dictioDependents = getDependentsAOMap();
    dictioHyphen = getHyphenAOMap();
    dictioObliquePron = getObliquePronounAOMap();
  }

  var body = DocumentApp.getActiveDocument().getBody();
  body.editAsText().replaceText("“", "\"");
  body.editAsText().replaceText("”", "\"");
  body.editAsText().replaceText("‘", "'");
  body.editAsText().replaceText("’", "'");
  var words = body.getText().split(/\r?\n| /);

  for (var i = 0; i < words.length; i++) {
    var word = removePunctuation(words[i]);

    if (word.includes('-')) {
      var suggestionsHyphen = dictioHyphen.get(word.toLowerCase());

      if (suggestionsHyphen) {
        mountChangebleObject(word, suggestionsHyphen, changebles, body);

      } else {
        var wordsHyphen = word.split('-');

        var lastPart = wordsHyphen[wordsHyphen.length-1];
        var suggestionsOblique = dictioObliquePron.get(lastPart.toLowerCase());

        if (suggestionsOblique) {
          var completeSuggestionsOblique = new Array();
          for (var x = 0; x < suggestionsOblique.length; x++) {
            if (suggestionsOblique[x] == '') continue;
            var wordSugg = word.substring(0, word.lastIndexOf('-') + 1);
            completeSuggestionsOblique.push(wordSugg.concat(suggestionsOblique[x]));
          }
          mountChangebleObject(word, completeSuggestionsOblique, changebles, body);

        } else {
          for (var y = 0; y < wordsHyphen.length; y++) {
            var suggestions = dictio.get(wordsHyphen[y].toLowerCase());

            if (suggestions) {
              mountChangebleObject(wordsHyphen[y], suggestions, changebles, body);
            }
          }
        }
      }

    } else {
      var suggestionsDep = dictioDependents.get(word.toLowerCase());

      if (suggestionsDep && !(i + 1 > words.length) && !punctuationAtLastChar(words[i])) {
        var nextWord = removePunctuation(words[i+1]);

        if (dictioNeutral.includes(nextWord)) {
          Logger.log('getSuggestions: suggestionsDep: nextWord to change found: ' + nextWord);

          var neutralSuggestion = new Array();
          neutralSuggestion.push(suggestionsDep[0] + " " + nextWord);
          mountChangebleObject(word + " " + nextWord, neutralSuggestion, changebles, body);
          i++;

        } else {
          var suggestions = dictio.get(nextWord.toLowerCase());

          if (suggestions) {
            var expression = word + " " + nextWord;
            Logger.log('getSuggestions: suggestionsDep: suggestions: expression to change found: ' + expression);

            var expressionSuggestions = new Array();

            for (var x = 0; x < suggestions.length; x++) {
              if (suggestions[x] == '') continue;

              expressionSuggestions.push(suggestionsDep[x] + " " + suggestions[x]);
            }

            mountChangebleObject(expression, expressionSuggestions, changebles, body);

            i++;
          }
        }
        
      } else {
        var suggestions = dictio.get(word.toLowerCase());

        if (suggestions) {
          mountChangebleObject(word, suggestions, changebles, body);
        }
      }
    }
  }

  CacheService.getDocumentCache().put("changebles", JSON.stringify(changebles));
  CacheService.getDocumentCache().put("totalLength", JSON.stringify(body.getText().length));
  CacheService.getDocumentCache().put("foundWord", null);

  return changebles;
}

function mountChangebleObject(word, suggestions, changebles, body) {
  var wordIndex = 0;

  for (var ix = 0; ix < changebles.length; ix++) {
    if (changebles[ix].word.localeCompare(word) == 0) {
      wordIndex++;
    }
  }

  var wordLocation = getWordLocation(word, wordIndex);
  Logger.log('getSuggestions: word to change found: ' + word + ' / wordIndex: ' + wordIndex);
  Logger.log(' / startIndex: ' + wordLocation.getStartOffset() + ' / endIndex: ' + wordLocation.getEndOffsetInclusive());

  setHighlight(wordLocation, true);
  changebles.push(new Changeble(word, suggestions, wordIndex, wordIndex));
}

function acceptSuggestion(e) {
  try {
    if (wasExternallyChanged()) {
      return showExternallyChangedErrorCard();
    }

    var foundWord = CacheService.getDocumentCache().get("foundWord");
    if (foundWord != null) {
      setHighlight(getWordLocation(foundWord, CacheService.getDocumentCache().get("foundWordIdx")), true);
      CacheService.getDocumentCache().put("foundWord", null);
    }

    var gridParam = e.parameters.grid_item_identifier;
    var params = gridParam.split("_");
    var word = params[0];
    var suggestion = params[1];
    var documentOccurIdx = params[2];
    var changebleOccurIdx = params[3];
    var wordLocation = getWordLocation(word, documentOccurIdx);
    var updateDocumentOccurIdx = !suggestion.endsWith("(a)") && !suggestion.endsWith("(as)") 
          && !suggestion.endsWith("/a") && !suggestion.endsWith("/as");
    
    setHighlight(wordLocation, false);
    // setCursorToWordPosition(wordLocation);

    var startIndex = getCorrectStartOffset(wordLocation);
    var endIndex = getCorrectEndOffset(wordLocation);
    var textElement = wordLocation.getElement().asText();
    if (startIndex == 0) {
      textElement.deleteText(startIndex+1, endIndex);
      textElement.insertText(startIndex+1, suggestion);
      textElement.deleteText(startIndex, startIndex);
    } else {
      textElement.deleteText(startIndex, endIndex);
      textElement.insertText(startIndex, suggestion);
    }

    updateChangebleList(word, changebleOccurIdx, updateDocumentOccurIdx);

    CacheService.getDocumentCache().put("totalLength", 
            JSON.stringify(DocumentApp.getActiveDocument().getBody().getText().length));
    
    return updateCard();

  } catch(e) {
    Logger.log(e);
    return showGeneralErrorCard();
  }
}

function ignoreSuggestion(e) { 
  try {
    if (wasExternallyChanged()) {
      return showExternallyChangedErrorCard();
    }

    var foundWord = CacheService.getDocumentCache().get("foundWord");
    if (foundWord != null) {
      setHighlight(getWordLocation(foundWord, CacheService.getDocumentCache().get("foundWordIdx")), true);
      CacheService.getDocumentCache().put("foundWord", null);
    }
    
    var word = e.parameters.word;
    var documentOccurIdx = e.parameters.documentOccurIdx;
    var changebleOccurIdx = e.parameters.changebleOccurIdx;
    var wordLocation = getWordLocation(word, documentOccurIdx);

    setHighlight(wordLocation, false);
    // setCursorToWordPosition(wordLocation);
    updateChangebleList(word, changebleOccurIdx, false);

    CacheService.getDocumentCache().put("totalLength", 
            JSON.stringify(DocumentApp.getActiveDocument().getBody().getText().length));

    return updateCard();

  } catch(e) {
    Logger.log(e);
    return showGeneralErrorCard();
  }
}

function findChangebleInDoc(e) {
  try {
    if (wasExternallyChanged()) {
      return showExternallyChangedErrorCard();
    }

    var foundWord = CacheService.getDocumentCache().get("foundWord");
    if (foundWord != null) {
      setHighlight(getWordLocation(foundWord, CacheService.getDocumentCache().get("foundWordIdx")), true);
    }
    
    var word = e.parameters.word;
    var documentOccurIdx = e.parameters.documentOccurIdx;
    var wordLocation = getWordLocation(word, documentOccurIdx);

    setCursorToWordPosition(wordLocation);
    setFindHighlight(wordLocation);

    CacheService.getDocumentCache().put("foundWord", word);
    CacheService.getDocumentCache().put("foundWordIdx", documentOccurIdx);

  } catch(e) {
    Logger.log(e);
    return showGeneralErrorCard();
  }
}

function Changeble(word, suggestions, documentOccurIdx, changebleOccurIdx) {
  this.word = word;
  this.suggestions = suggestions;
  this.documentOccurIdx = documentOccurIdx;
  this.changebleOccurIdx = changebleOccurIdx;
}

function setHighlight(wordLocation, enable) {
  var style = {};
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = enable ? "#FFFF00" : "#FFFFFF";

  wordLocation.getElement().setAttributes(getCorrectStartOffset(wordLocation), getCorrectEndOffset(wordLocation), style);
}

function setFindHighlight(wordLocation) {
  var style = {};
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = "#FF9933";

  wordLocation.getElement().setAttributes(getCorrectStartOffset(wordLocation), getCorrectEndOffset(wordLocation), style);
}

function getWordLocation(word, documentOccurIdx) {
  var body = DocumentApp.getActiveDocument().getBody();
  var wordLocation = body.findText(leftBoundary + word + rightBoundary, wordLocation);
  var index = 0;

  while (index != documentOccurIdx) {
    wordLocation = body.findText(leftBoundary + word + rightBoundary, wordLocation);
    index++;
  }

  return wordLocation;
}

function getCorrectStartOffset(wordLocation) {
  var word = wordLocation.getElement().asText().getText();
  word = word.substring(wordLocation.getStartOffset(), wordLocation.getEndOffsetInclusive());
  
  if (word.substring(0, 1).search(/\s|\d|[!"“”#$%&\'()*+,\-.:;<=>?@[\]^_`{|}~]/g) != -1) {
    return parseInt(wordLocation.getStartOffset()) + 1;
  }
  return wordLocation.getStartOffset();
}

function getCorrectEndOffset(wordLocation) {
  var word = wordLocation.getElement().asText().getText();
  word = word.substring(wordLocation.getStartOffset(), wordLocation.getEndOffsetInclusive() + 1);
  
  if (word.substring(word.length-1).search(/\s|\d|[!"“”#$%&\'()*+,\-.:;<=>?@[\]^_`{|}~]/g) != -1) {
    return parseInt(wordLocation.getEndOffsetInclusive()) - 1;
  }
  return wordLocation.getEndOffsetInclusive();
}

function punctuationAtLastChar(word) {
  var punct = /[\u2000-\u206F\u2E00-\u2E7F\\'!"“”#$%&*+,.:;<=>?@\[\]^_`{|}~]/g;
  return word.substring(word.length-1).search(punct) != -1;
}

function removePunctuation(word) {
  var punct = /[\u2000-\u206F\u2E00-\u2E7F\\'!"“”#$%&*+,.:;<=>?@\[\]^_`{|}~]/g;
  word = word.replaceAll(punct, '');

  if (word.startsWith('(')) {
    word = word.substring(1, word.length);
  }
  if (word.endsWith(')') && !word.endsWith('(a)') && !word.endsWith("(as)")) {
    word = word.substring(0, word.length -1);
  }

  return word;
}

function validateDocAndSetCursorToStart() {

  var doc = DocumentApp.getActiveDocument();
  var paragraph = doc.getBody().getChild(0);

  if(paragraph.getNumChildren() == 0) {
    return showOpeningDocErrorCard();
  }
  
  var position = doc.newPosition(paragraph.getChild(0), 0);
  
  doc.setCursor(position);
}

function setCursorToWordPosition(wordLocation) {
  var wordPosition = DocumentApp.getActiveDocument().newPosition(wordLocation.getElement(), 0);
  DocumentApp.getActiveDocument().setCursor(wordPosition);
}

function updateChangebleList(word, changebleOccurIdx, updateDocumentOccurIdx) {
  var changeblesTmp = JSON.parse(CacheService.getDocumentCache().get("changebles"));
  var changebles = new Array();

  var index = 0;
  for (var i = 0; i < changeblesTmp.length; i++) {
    if (changeblesTmp[i].word.localeCompare(word) == 0) {
      Logger.log("same word: " + word + " / index: " + index);
      if (changeblesTmp[i].changebleOccurIdx != changebleOccurIdx) {
        var changeble = updateDocumentOccurIdx ? 
                          new Changeble(changeblesTmp[i].word, changeblesTmp[i].suggestions, index, index)
                          : new Changeble(changeblesTmp[i].word, changeblesTmp[i].suggestions, changeblesTmp[i].documentOccurIdx, index);
        changebles.push(changeble);
        index++;
      }
    } else {
      Logger.log("not the same: " + changeblesTmp[i].word + " / " + word);
      changebles.push(changeblesTmp[i]);
    }
  }
  changebles.forEach(changeble => Logger.log('changeble: ' + JSON.stringify(changeble)));
  
  CacheService.getDocumentCache().put("changebles", JSON.stringify(changebles));
}

function wasExternallyChanged() {
  var totalLength = JSON.parse(CacheService.getDocumentCache().get("totalLength"));
  return totalLength != DocumentApp.getActiveDocument().getBody().getText().length;
}

function showExternallyChangedErrorCard() {
  var builder = CardService.newCardBuilder().addSection(CardService.newCardSection()
      .addWidget(CardService
      .newTextParagraph()
      .setText("<b>Infelizmente, ainda não lidamos com alterações externas durante o uso deste complemento.<br><br>Poderia terminar as alterações e reiniciar depois?</b>")));

    return CardService.newNavigation().updateCard(builder.build());
}

function showGeneralErrorCard() {
  var builder = CardService.newCardBuilder().addSection(CardService.newCardSection()
      .addWidget(CardService
      .newTextParagraph()
      .setText("<b>Ops, desculpa! Algo deu errado!<br><br>Poderia atualizar o complemento e recomeçar?</b>")));

    return builder.build();
}

function showOpeningDocErrorCard() {
  var builder = CardService.newCardBuilder().addSection(CardService.newCardSection()
      .addWidget(CardService
      .newTextParagraph()
      .setText("<b>Ops!<br><br>Este documento não é suportado por não ter as permissões necessárias ou por não ser um documento original Google Docs.</b>")));

    return builder.build();
}
