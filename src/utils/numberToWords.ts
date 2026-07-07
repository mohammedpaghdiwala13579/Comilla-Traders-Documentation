// Helper to convert number to words in Indian numbering system (Lakhs, Crores) with "Taka Only" appended at the end
export function numberToWords(num: number): string {
  if (num === 0) return "Zero Taka Only";

  const ones = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen"
  ];
  const tens = [
    "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"
  ];

  const convertTwoDigits = (n: number): string => {
    if (n === 0) return "";
    if (n < 20) return ones[n];
    const digit = n % 10;
    const ten = Math.floor(n / 10);
    return tens[ten] + (digit ? " " + ones[digit] : "");
  };

  const convertThreeDigits = (n: number): string => {
    const hundred = Math.floor(n / 100);
    const rem = n % 100;
    let word = "";
    if (hundred > 0) {
      word = ones[hundred] + " Hundred";
    }
    if (rem > 0) {
      word += (word ? " and " : "") + convertTwoDigits(rem);
    }
    return word;
  };

  const convertIndian = (n: number): string => {
    if (n === 0) return "";
    
    let words = "";
    
    // Crores (1,00,00,000)
    const crores = Math.floor(n / 10000000);
    let rem = n % 10000000;
    if (crores > 0) {
      words += (crores >= 100 ? convertIndian(crores) : convertTwoDigits(crores)) + " Crore ";
    }
    
    // Lakhs (1,00,00,000) -> 1,00,000 is 1 Lakh
    const lakhs = Math.floor(rem / 100000);
    rem = rem % 100000;
    if (lakhs > 0) {
      words += convertTwoDigits(lakhs) + " Lakh ";
    }
    
    // Thousands (1,000)
    const thousands = Math.floor(rem / 1000);
    rem = rem % 1000;
    if (thousands > 0) {
      words += convertTwoDigits(thousands) + " Thousand ";
    }
    
    // Remaining less than 1,000
    if (rem > 0) {
      words += convertThreeDigits(rem);
    }
    
    return words.trim();
  };

  const fixedStr = num.toFixed(2);
  const [integerStr, decimalStr] = fixedStr.split(".");
  const integerPart = parseInt(integerStr, 10);
  const decimalPart = parseInt(decimalStr, 10);

  let result = "";
  if (integerPart === 0) {
    result = "Zero";
  } else {
    result = convertIndian(integerPart);
  }

  if (decimalPart > 0) {
    result += " and " + convertTwoDigits(decimalPart) + " Paisa";
  }

  return result.trim() + " Taka Only";
}
