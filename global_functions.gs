function searchValueRange(targetValue, range, offsetRow=0, offsetColumn=0) {
  if (!targetValue || !range) return;
  let returnValue = new Array();

  try{
    const value = range.createTextFinder(targetValue).matchEntireCell(true).findAll();
    const valueRange = value.map(range => range);

    for(let i = 0; i < valueRange.length; i++) {
      let offsetValueRange = valueRange[i].offset(offsetRow,offsetColumn);
      returnValue.push(offsetValueRange);
      console.log('valueRange[' + i + ']: ' + targetValue + ' ' + offsetValueRange.getA1Notation() + ' offset(' + offsetRow + ',' + offsetColumn + ')');
    }
  } catch(error) {
    console.log('Continue error: ' + error.message);
  }

  return returnValue;
}

function katakanaToHiragana(input) {
  return input.replace(/[\u30A1-\u30F6]/g, function(match) {
    return String.fromCharCode(match.charCodeAt(0) - 0x60);
  });
}