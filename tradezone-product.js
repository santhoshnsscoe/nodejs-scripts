const { readExcelFile, writeExcelFile } = require("./utils/excel");
const { handlize, nl2br, createKey } = require("./utils/default");

// constants
const INPUT_FILE_PATH = "./files/tradezone-products.csv";
const OUTPUT_FILE_PATH = "./files/tradezone-products-updated.csv";
const OUTPUT_SKIPPED_FILE_PATH = "./files/tradezone-products-skipped.csv";
const OUTPUT_NO_MARKUP_FILE_PATH = "./files/tradezone-products-no-markup.csv";

const PRODUCTS_FILE_PATH = "./files/tradezone-products.xlsx";
const IMAGES_FILE_PATH = "./files/tradezone-images.csv";
const MARKUPS_FILE_PATH = "./files/tradezone-markup.xlsx";

// updated products
const updatedProducts = [];
const skippedProducts = [];
const noMarkupProducts = [];
const productsData = new Map();
const markupsData = new Map();
const imagesData = new Map();

/**
 * Get the body HTML for the product
 * @param {Object} product
 * @returns {string}
 */
const getBodyHtml = (product) => {
  return `
    <p>${nl2br(product["Product Details"])}</p>
    <p>Warranty Information:</p>
    <p>${nl2br(product["Warranty Information (From Install"])}</p>
    <p>Attributes:</p>
    <p>${product["Attributes"]}</p>
    <p>Shipping Information:</p>
    <p>${product["Shipping Information"]}</p>
  `;
};

/**
 * Extract the shipping data from the text
 * @param {string} text
 * @returns {Object}
 */
const extractShippingData = (text) => {
  const clean = text.replace(/\r/g, "").trim();

  // Helper: extract first matching group
  const match = (regex) => {
    const m = clean.match(regex);
    return m ? m[1].trim() : "";
  };

  // Detect weight + unit automatically: supports kg, g, lbs etc.
  const weightRegex = /Weight\s*\((.*?)\)\s*([\d\.]+)/i;
  const weightMatch = clean.match(weightRegex);

  const weight = weightMatch ? weightMatch[2] : "";
  const weightUnit = weightMatch ? weightMatch[1].toLowerCase() : "";

  return {
    // Weight
    weight,
    weight_unit: weightUnit,

    // Main dimensions
    length: match(/Length\s*\(mm\)\s*([\d\.]+)/i),
    height: match(/Height\s*\(mm\)\s*([\d\.]+)/i),
    width: match(/Width\s*\(mm\)\s*([\d\.]+)/i),

    // Packaging dimensions
    packaging_length: match(/Length Packaging\s*\(mm\)\s*([\d\.]+)/i),
    packaging_height: match(/Height Packaging\s*\(mm\)\s*([\d\.]+)/i),
    packaging_width: match(/Width Packaging\s*\(mm\)\s*([\d\.]+)/i),

    // Barcodes
    barcode: match(/Barcode\s*\n\s*([\d]+)/i),
    barcode_secondary: match(/Barcode \(Secondary\)\s*\n\s*([\d]+)/i),
    barcode_tertiary: match(/Barcode \(Tertiary\)\s*\n\s*([\d]+)/i),
  };
};

/**
 * Get product data
 */
const addProductData = ({
  product,
  title,
  handle,
  tradezonePartNumber,
  supplierPartNumber,
  type,
  bodyHtml,
  price,
  cost,
  shippingData,
  productData,
  skippedProduct,
  noMarkupFound,
  images,
}) => {
  const imageSrc = images[0] || (productData && productData["Image1"]) || "";
  const imagePosition = imageSrc ? 1 : "";
  const importData = {
    Handle: handle,
    Title: title,
    "Body (HTML)": bodyHtml,
    Vendor: product["Manufacturer"] || "All Led Direct",
    Type: type,
    Tags: product["Search Terms"],
    Status: "active",
    "Option1 Name": "Title",
    "Option1 Value": "Default Title",
    "Variant SKU": "",
    "Variant Grams": shippingData.weight,
    "Variant Weight Unit": shippingData.weight_unit,
    "Variant Inventory Tracker": "shopify",
    "Variant Inventory Qty": "0",
    "Variant Inventory Policy": "deny",
    "Variant Fulfillment Service": "manual",
    "Variant Price": price,
    "Cost per item": cost,
    "Variant Requires Shipping": true,
    "Variant Taxable": true,
    "Variant Barcode": shippingData.barcode,
    "Image Src": imageSrc,
    "Image Position": imagePosition,
    "Image Alt Text": "",
    "Tradezone Part Number (product.metafields.tradezone.part_number)":
      tradezonePartNumber,
    "Supplier Part Number (product.metafields.tradezone.supplier_part_number)":
      supplierPartNumber,
    "Sub Group (product.metafields.tradezone.sub_group)": product["Sub Group"],
    "Warranty Information (product.metafields.tradezone.warranty)":
      product["Warranty Information (From Install"],
    "Attributes (product.metafields.tradezone.attributes)":
      product["Attributes"],
    "Shipping Information (product.metafields.tradezone.shipping)":
      product["Shipping Information"],
    "Length (product.metafields.tradezone.length)": shippingData.length,
    "Height (product.metafields.tradezone.height)": shippingData.height,
    "Width (product.metafields.tradezone.width)": shippingData.width,
    "Length Packaging (product.metafields.tradezone.length_packaging)":
      shippingData.packaging_length,
    "Height Packaging (product.metafields.tradezone.height_packaging)":
      shippingData.packaging_height,
    "Width Packaging (product.metafields.tradezone.width_packaging)":
      shippingData.packaging_width,
    "Barcode (product.metafields.tradezone.barcode)": shippingData.barcode,
  };

  if (skippedProduct) {
    skippedProducts.push(importData);
  } else if (noMarkupFound) {
    noMarkupProducts.push(importData);
  } else {
    updatedProducts.push(importData);
  }
};

/**
 * Get the image data
 * @param {Object} product
 * @returns {Object}
 */
const addImageData = ({
  handle,
  imageSrc,
  imagePosition,
  skippedProduct,
  noMarkupFound,
}) => {
  const importData = {
    Handle: handle,
    Title: "",
    "Body (HTML)": "",
    Vendor: "",
    Type: "",
    Tags: "",
    Status: "",
    "Option1 Name": "",
    "Option1 Value": " ",
    "Variant SKU": "",
    "Variant Grams": "",
    "Variant Weight Unit": "",
    "Variant Inventory Tracker": "",
    "Variant Inventory Qty": "",
    "Variant Inventory Policy": "",
    "Variant Fulfillment Service": "",
    "Variant Price": "",
    "Cost per item": "",
    "Variant Requires Shipping": "",
    "Variant Taxable": "",
    "Variant Barcode": "",
    "Image Src": imageSrc,
    "Image Position": imagePosition,
    "Image Alt Text": "",
    "Tradezone Part Number (product.metafields.tradezone.part_number)": "",
    "Supplier Part Number (product.metafields.tradezone.supplier_part_number)":
      "",
    "Sub Group (product.metafields.tradezone.sub_group)": "",
    "Length (product.metafields.tradezone.length)": "",
    "Height (product.metafields.tradezone.height)": "",
    "Width (product.metafields.tradezone.width)": "",
    "Length Packaging (product.metafields.tradezone.length_packaging)": "",
    "Height Packaging (product.metafields.tradezone.height_packaging)": "",
    "Width Packaging (product.metafields.tradezone.width_packaging)": "",
    "Barcode (product.metafields.tradezone.barcode)": "",
  };
  if (skippedProduct) {
    skippedProducts.push(importData);
  } else if (noMarkupFound) {
    noMarkupProducts.push(importData);
  } else {
    updatedProducts.push(importData);
  }
};

/**
 * Set the products data
 */
const setProductsData = () => {
  // read the input file
  const products = readExcelFile(PRODUCTS_FILE_PATH);
  for (const product of products) {
    productsData.set(`t-${createKey(product["Product Title"])}`, product);
    productsData.set(`s-${createKey(product["Part number"])}`, product);
    productsData.set(`p-${createKey(product["SKU"])}`, product);
  }
};

/**
 * Set the markups data
 */
const setMarkupsData = () => {
  // read the input file
  const markups = readExcelFile(MARKUPS_FILE_PATH);
  let mainCategory = "";
  for (const markup of markups) {
    // get the category
    const category = markup["Category"] || "";
    if (!category) {
      continue;
    }

    // if the website list price mark up % is set, add the markup to the markups data
    if (markup["Website list price mark up %"]) {
      let key = createKey(category);
      if (markupsData.has(key)) {
        key = `${key}-${createKey(mainCategory)}`;
      }
      markupsData.set(key, {
        main: mainCategory,
        category: markup["Category"],
        markup: 1 + Number(markup["Website list price mark up %"]),
      });
    } else {
      mainCategory = markup["Category"];
    }
  }
};

/**
 * Set the images data
 */
const setImagesData = () => {
  // read the input file
  const images = readExcelFile(IMAGES_FILE_PATH);
  for (const image of images) {
    const key = createKey(image["Title"]);
    const imageSrc = image["Image Src"];
    if (!imageSrc) {
      continue;
    }
    const images = imagesData.get(key) || [];
    images.push(imageSrc);
    imagesData.set(key, images);
  }
};

/**
 * Update the products
 */
const updateProducts = () => {
  // initialize the products data
  setProductsData();
  setMarkupsData();
  setImagesData();

  // read the input file
  const products = readExcelFile(INPUT_FILE_PATH);
  let noProductDataFoundCount = 0;
  let updatedProductsCount = 0;
  let skippedProductsCount = 0;
  let noMarkupProductsCount = 0;

  /**
   * mapping
   * Description = Title
   * Group = Type
   * Supplier Part Number = Variant SKU
   * Tradezone Part Number = Tradezone Part Number (product.metafields.tradezone.part_number)
   * Cost Price = Cost per item
   * Search Terms = Tags
   * Sub Group = Sub Group (product.metafields.tradezone.sub_group)
   * Product Details + Warranty Information (From Install) + Shipping Information + Attributes = Body (HTML)
   */

  for (const product of products) {
    let skippedProduct = false;
    let noMarkupFound = false;

    // get basic values from the product
    const title = product["Description"];
    const type = product["Group"];
    const supplierPartNumber = product["Supplier Part Number"];
    const tradezonePartNumber = product["Tradezone Part Number"];
    let markupAmount = 2;

    // get the cost and skip if it's 0 or less
    const cost = Number(product["Cost Price"] || "0");
    if (cost <= 0) {
      //console.log("Skipping product due zero cost: ", title);
      skippedProduct = true;
      skippedProductsCount++;
    }

    // get the product data from the products data
    let productData = productsData.get(`t-${createKey(title)}`);
    if (!productData) {
      productData = productsData.get(`p-${createKey(tradezonePartNumber)}`);
    }
    if (!productData) {
      productData = productsData.get(`s-${createKey(supplierPartNumber)}`);
    }
    if (!productData) {
      noProductDataFoundCount++;
    }

    const markup = markupsData.get(createKey(type));
    if (!markup) {
      //console.log("Skipping product due no markup: ", title);
      noMarkupFound = true;
      noMarkupProductsCount++;
    } else {
      markupAmount = markup.markup;
    }

    // get the images data
    const images = imagesData.get(createKey(title)) || [];

    // get the shipping data
    const shippingData = extractShippingData(product["Shipping Information"]);
    const weight = Number(shippingData.weight || "0");
    if (weight <= 0) {
      //console.log("Skipping product due zero weight: ", title);
      skippedProduct = true;
      skippedProductsCount++;
    }

    // get derived values from the product
    const handle = handlize(title);
    const bodyHtml = getBodyHtml(product);
    const price = cost * markupAmount;

    if (skippedProduct === false && noMarkupFound === false) {
      updatedProductsCount++;
    }

    // add the product to the updated products
    addProductData({
      product,
      title,
      handle,
      tradezonePartNumber,
      supplierPartNumber,
      type,
      bodyHtml,
      price,
      cost,
      shippingData,
      productData,
      skippedProduct,
      noMarkupFound,
      images,
    });

    if (images.length > 1) {
      addImageData({
        handle,
        imageSrc: images[1],
        imagePosition: 2,
        skippedProduct,
        noMarkupFound,
      });
    } else if (productData && productData["Image2"]) {
      addImageData({
        handle,
        imageSrc: productData["Image2"],
        imagePosition: 2,
        skippedProduct,
        noMarkupFound,
      });
    }

    if (images.length > 2) {
      addImageData({
        handle,
        imageSrc: images[2],
        imagePosition: 3,
        skippedProduct,
        noMarkupFound,
      });
    } else if (productData && productData["Image3"]) {
      addImageData({
        handle,
        imageSrc: productData["Image3"],
        imagePosition: 3,
        skippedProduct,
        noMarkupFound,
      });
    }

    if (images.length > 3) {
      addImageData({
        handle,
        imageSrc: images[3],
        imagePosition: 4,
        skippedProduct,
        noMarkupFound,
      });
    } else if (productData && productData["Image4"]) {
      addImageData({
        handle,
        imageSrc: productData["Image4"],
        imagePosition: 4,
        skippedProduct,
        noMarkupFound,
      });
    }

    if (images.length > 4) {
      addImageData({
        handle,
        imageSrc: images[4],
        imagePosition: 5,
        skippedProduct,
        noMarkupFound,
      });
    } else if (productData && productData["Image5"]) {
      addImageData({
        handle,
        imageSrc: productData["Image5"],
        imagePosition: 5,
        skippedProduct,
        noMarkupFound,
      });
    }
  }

  // log the skipped products
  console.log("Skipped products: ", skippedProductsCount);
  console.log("No markup found: ", noMarkupProductsCount);
  console.log("Updated products: ", updatedProductsCount);
  console.log("No product data found: ", noProductDataFoundCount);

  // write the updated products to the output file
  writeExcelFile(OUTPUT_FILE_PATH, updatedProducts);
  writeExcelFile(OUTPUT_SKIPPED_FILE_PATH, skippedProducts);
  writeExcelFile(OUTPUT_NO_MARKUP_FILE_PATH, noMarkupProducts);
};

// run the script
updateProducts();
