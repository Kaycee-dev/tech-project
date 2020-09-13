function unique(arr1) {
  let result = {};
  arr1.forEach(item => {
    switch (result.hasOwnProperty(item)) {
      case true:
        result[item] += 1;
        break;
      default:
        result[item] = 1;
        break;
    }
  });

  return result;
}

console.log(unique(['dog', 'cat', 'sheep', 'cat', 'sheep']));