function invest (amountToInvest, numberOfYearsInvesting){
	if (numberOfYearsInvesting < 3) {
		return `Your interest with us after ${numberOfYearsInvesting} years is ${Number.parseFloat(amountToInvest * 0.015 * numberOfYearsInvesting).toFixed(2)} and total payable amount is ${Math.round((amountToInvest * (1 + (0.015 * numberOfYearsInvesting))*100)/100).toFixed(2)}`
	}else if (numberOfYearsInvesting >= 3 && numberOfYearsInvesting <= 5) {
		return `Your interest with us after ${numberOfYearsInvesting} years is ${Number.parseFloat(amountToInvest * 0.025 * numberOfYearsInvesting).toFixed(2)} and total payable amount is ${Math.round((amountToInvest * (1 + (0.025 * numberOfYearsInvesting))*100)/100).toFixed(2)}`
	}else if (numberOfYearsInvesting < 10) {
		return `Your interest with us after ${numberOfYearsInvesting} years is ${Number.parseFloat(amountToInvest * 0.035 * numberOfYearsInvesting).toFixed(2)} and total payable amount is ${Math.round((amountToInvest * (1 + (0.035 * numberOfYearsInvesting))*100)/100).toFixed(2)}`
	}else {
		return `Your interest with us after ${numberOfYearsInvesting} years is ${Number.parseFloat(amountToInvest * 0.05 * numberOfYearsInvesting).toFixed(2)} and total payable amount is ${Math.round((amountToInvest * (1 + (0.05 * numberOfYearsInvesting))*100)/100).toFixed(2)}`
	}
	
}

console.log(invest(10000,1));