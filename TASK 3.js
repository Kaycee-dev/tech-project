function testArray(arr1, ...arrN) {
  let finalArray = [];

  if (typeof arr1 === 'undefined'){
    
  }else if (typeof arr1 !== 'object') {
    finalArray.push(arr1);
    unpackArr(arrN);
  }else if (
    (typeof arr1 === 'object' && arr1 != null && arr1.length > 0) ||
    arrN.length > 0
  ) {
    unpackArr(arr1);
    unpackArr(arrN);
  }

  return finalArray
    .sort(function (a, b) {
      return a - b;
    })
    .join(',');

  function unpackArr(a) {
    a.forEach(item => {
      if (typeof item == 'object') {
        unpackArr(item);
      } else {
        finalArray.push(item);
      }
    });
  }
}

console.log(testArray([9,8,4,1],[6,3,4,0,-1], [ ] , [4,2,7,4,1,0]));
