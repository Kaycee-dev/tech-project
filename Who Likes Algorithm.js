function likes(arr) {
	let lenArr = arr.length;
	let reply = '';
	switch (lenArr) {
		case 0:
			reply = 'no one likes this item';
			break;
		case 1:
			reply = `${arr[0]} likes this item`;
			break;
		case 2:
			reply = `${arr[0]} and ${arr[1]} like this item`;
			break;
		case 3:
			reply = `${arr[0]}, ${arr[1]} and ${arr[2]} like this item`
			break;
		default:
			reply = `${arr[0]}, ${arr[1]} and ${lenArr - 2} others like this item`
	}

	return reply

}

console.log(likes(["Soji", "Samuel", "Jane", "Kelechi"]))