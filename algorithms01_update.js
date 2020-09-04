
function invest (amountToInvest, numberOfYearsInvesting){
let years = numberOfYearsInvesting > 1 ? 'years':'year'
	if (numberOfYearsInvesting < 3) {
		return `Your interest with us after ${numberOfYearsInvesting} ${years} is ${Math.round(amountToInvest * 0.015 * numberOfYearsInvesting)} and total payable amount is ${Math.round((amountToInvest * (1 + (0.015 * numberOfYearsInvesting))*100)/100)}`
	}else if (numberOfYearsInvesting >= 3 && numberOfYearsInvesting <= 5) {
		return `Your interest with us after ${numberOfYearsInvesting} ${years} is ${Math.round(amountToInvest * 0.025 * numberOfYearsInvesting)} and total payable amount is ${Math.round((amountToInvest * (1 + (0.025 * numberOfYearsInvesting))*100)/100)}`
	}else if (numberOfYearsInvesting < 10) {
		return `Your interest with us after ${numberOfYearsInvesting} ${years} is ${Math.round(amountToInvest * 0.035 * numberOfYearsInvesting)} and total payable amount is ${Math.round((amountToInvest * (1 + (0.035 * numberOfYearsInvesting))*100)/100)}`
	}else {
		return `Your interest with us after ${numberOfYearsInvesting} ${years} is ${Math.round(amountToInvest * 0.05 * numberOfYearsInvesting)} and total payable amount is ${Math.round((amountToInvest * (1 + (0.05 * numberOfYearsInvesting))*100)/100)}`
	}
	
}

console.log(invest(10000,1));