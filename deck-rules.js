(function (global) {
  function isBasicLand(card) {
    if (!card?.typeLine && !card?.cardType) {
      return false;
    }

    return String(card.typeLine || card.cardType || '').toLowerCase().includes('basic land');
  }

  function getCardOracleTexts(card) {
    const texts = [String(card?.oracleText || '')];
    if (Array.isArray(card?.cardFaces)) {
      card.cardFaces.forEach((face) => texts.push(String(face?.oracleText || '')));
    }
    return texts;
  }

  function isUnlimitedCopiesCard(card) {
    return getCardOracleTexts(card).some((text) => /a deck can have any number of cards named/i.test(text));
  }

  function allowsMultipleCopies(card) {
    return isBasicLand(card) || isUnlimitedCopiesCard(card);
  }

  global.CommanderDeckRules = Object.freeze({
    isBasicLand,
    getCardOracleTexts,
    isUnlimitedCopiesCard,
    allowsMultipleCopies,
  });
})(window);
