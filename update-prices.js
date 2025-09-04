const { readExcelFile, writeExcelFile } = require("./utils/excel");

// constants
const INPUT_FILE_PATH = "./files/products-02.xlsx";
const OUTPUT_FILE_PATH = "./files/products-02-updated.xlsx";
const SHEET_NAME = "Products";
const PRICE_PERCENTAGE = 12;

// updated products
let updatedProducts = [];

/**
 * Update the prices of the products
 */
const updatePrices = () => {
  // read the input file
  const products = readExcelFile(INPUT_FILE_PATH);
  
  for (const product of products) {
    let variantPrice = Number(product["Variant Price"] || 0);
    let variantCompareAtPrice = Number(product["Variant Compare At Price"] || 0);

    if (variantPrice < variantCompareAtPrice) {
      variantPrice = Math.ceil(variantCompareAtPrice - ((variantCompareAtPrice * PRICE_PERCENTAGE) / 100));
    } else {
      variantCompareAtPrice = variantPrice;
      variantPrice = Math.ceil(variantPrice - ((variantPrice * PRICE_PERCENTAGE) / 100));
    }
    updatedProducts.push({
      ...product,
      "Variant Price": variantPrice,
      "Variant Compare At Price": variantCompareAtPrice,
    });
  }

  // write the updated products to the output file
  writeExcelFile(OUTPUT_FILE_PATH, updatedProducts, SHEET_NAME);
};

// run the script
updatePrices();
