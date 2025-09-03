const { readExcelFile, writeExcelFile } = require("./utils/excel");

// constants
const INPUT_FILE_PATH = "./files/products-01.xlsx";
const OUTPUT_FILE_PATH = "./files/products-01-updated.xlsx";
const SHEET_NAME = "Products";

// Product types linked to collections
const PRODUCT_TYPES = [
  {
    code: "collection-handle-1",
    name: "Type Name 1",
  },
  {
    code: "collection-handle-2",
    name: "Type Name 2",
  },
];

// get collection names from the input file - testing
//let collections = [];

// updated products
let updatedProducts = [];

/**
 * Update the types of the products
 */
const updateTypes = () => {
  // read the input file
  const products = readExcelFile(INPUT_FILE_PATH);
  
  for (const product of products) {
    // get the custom collections from the product
    const customCollections = `${product["Custom Collections"]}`
      .split(",")
      .map((collection) => collection.trim());

    // get list of collections - testing
    //collections = collections.concat(customCollections);

    // get the type from the product types
    const type = PRODUCT_TYPES.find((type) =>
      customCollections.includes(type.code)
    );

    // if the type is different, log the product
    //if (type?.name != product.Type) {
    //  console.log(product.Type, "!= ", type?.name);
    //}

    // add the product to the updated products
    updatedProducts.push({
      ID: product.ID,
      Handle: product.Handle,
      Title: product.Title,
      Type: type?.name || product.Type,
      Collections: product["Custom Collections"],
    });
  }

  // get the unique collections - testing
  //collections = [...new Set(collections)];
  //console.log(collections);

  // write the updated products to the output file
  writeExcelFile(OUTPUT_FILE_PATH, updatedProducts, SHEET_NAME);
};

// run the script
updateTypes();
