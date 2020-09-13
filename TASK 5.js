function sum(arr1) {
  let result = 0;
  arr1.forEach(item => {
    switch (item % 6) {
      case 0:
        result += item;
        break;
    }
  })

  return result;
}


console.log(sum([2, 12, 18]));