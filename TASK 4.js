function matrixArray(arr1) {
  let finalArray,a1,a2,a3,b1,b2,b3,c1,c2,c3;
  finalArray = [];
  unpackArr(arr1);
  [a1,a2,a3,b1,b2,b3,c1,c2,c3] = finalArray;
  let determinant = a1*(b2*c3-c2*b3)-b1*(a2*c3-a3*c2)+c1*(a2*b3-a3*b2)
  return `Your determinant for ${arr1[0]}, ${arr1[1]} and ${arr1[2]} === ${determinant}`;

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

console.log(matrixArray([ 

[1 , 2 , 3] , 

[4 , 3, 2],

[2 , 1 , 0] 

]));