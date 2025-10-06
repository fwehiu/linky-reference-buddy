/**
 * Fruit Products Pricing Calculator - Web App Version
 * VERSION: WebApp-v12 / Script v23 (Fixed final calculation to include discounts)
 */

// --- Constants ---
const BULK_DISCOUNT_SHEET_NAME = 'Bulk Discounts';
const BASE_RATES_SHEET_NAME = 'Base Rates per Fruit Type';
const CAMERA_PRICE_CELL = 'E3';
const INFLATION_RATE_CELL = 'D3';
const MIN_PRICE_START_ROW = 6;
const DOC_TEMPLATE_ID = '1sUO5ivgJELWlY6aMEWZ1vI5chfQo9ONCbw3Wcef5p4I';

// --- Theme Colors ---
const THEME = { primary: '#212529', secondary: '#6c757d', light: '#f8f9fa', white: '#FFFFFF', red_accent: '#dc3545', border: '#dee2e6', adjustment: '#fd7e14', savingsColor: '#198754' };

// --- Helper function to round up to nearest hundred ---
function roundUpToNearestHundred(value) {
  return Math.ceil(value / 100) * 100;
}

// --- Web App Entry Point ---
function doGet(e) {
  Logger.log("Web App doGet triggered.");
  return HtmlService.createHtmlOutputFromFile('Calculator')
      .setTitle('Hectre Fruit Pricing Calculator')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- Functions needed for Spreadsheet Integration ---
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Fruit Calculator').addItem('Setup Calculator', 'setupCalculator').addItem('Launch Sidebar', 'showCalculatorSidebar').addToUi();
}

function setupCalculator() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calcSheetName = 'Calc';
  var calcSheet = ss.getSheetByName(calcSheetName);

  if (!calcSheet) {
    calcSheet = ss.insertSheet(calcSheetName);
  } else {
    calcSheet.clear();
  }

  calcSheet.setColumnWidth(1, 200)
           .setColumnWidth(2, 300)
           .setColumnWidth(3, 150);

  calcSheet.getRange('A1:C1')
           .merge()
           .setValue('Fruit Products Pricing Calculator')
           .setFontSize(18)
           .setFontWeight('bold')
           .setBackground(THEME.light)
           .setFontColor(THEME.primary)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle')
           .setBorder(true, true, true, true, true, true, THEME.white, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  calcSheet.setRowHeight(1, 40);
  calcSheet.getRange('B3').setValue('Company name');
  calcSheet.getRange('C3').setValue('Date');
  calcSheet.getRange('A1:C50').setBorder(null, null, null, null, null, null);

  calcSheet.getRange('A6:C6')
           .merge()
           .setValue('Results from Calculator:')
           .setFontWeight('bold')
           .setFontSize(14)
           .setBackground(THEME.light)
           .setFontColor(THEME.primary)
           .setBorder(null, null, true, null, null, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID_THIN);

  calcSheet.getRange('A8:C8')
           .setValues([['Item', 'Details', 'Price']])
           .setBackground(THEME.primary)
           .setFontColor(THEME.white)
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setBorder(null, null, true, null, null, true, THEME.primary, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  SpreadsheetApp.getUi().alert('Calculator sheet "' + calcSheetName + '" has been set up.');
}

function showCalculatorSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Calculator')
      .setTitle('Fruit Pricing Calculator')
      .setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- Data Retrieval Function ---
function getData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var baseRatesSheet = ss.getSheetByName(BASE_RATES_SHEET_NAME);
    var discountsSheet = ss.getSheetByName(BULK_DISCOUNT_SHEET_NAME);

    if (!baseRatesSheet || !discountsSheet) {
      throw new Error(`Required sheets not found.`);
    }

    var growerFruitTypesRaw = baseRatesSheet.getRange('A3:A' + baseRatesSheet.getLastRow()).getValues().flat();
    var packerFruitTypesRaw = baseRatesSheet.getRange('K3:K' + baseRatesSheet.getLastRow()).getValues().flat();
    var growerProductsRaw = baseRatesSheet.getRange('B2:I2').getValues().flat();
    var packerProductsRaw = baseRatesSheet.getRange('L2:N2').getValues().flat();

    const trimAndFilter = arr => arr.map(item => typeof item === 'string' ? item.trim() : item).filter(item => item !== null && item !== undefined && item !== '');

    var growerFruitTypes = trimAndFilter(growerFruitTypesRaw);
    var packerFruitTypes = trimAndFilter(packerFruitTypesRaw);
    var growerProducts = trimAndFilter(growerProductsRaw);
    var packerProducts = trimAndFilter(packerProductsRaw);

    var growerRates = baseRatesSheet.getRange(3, 2, growerFruitTypes.length || 1, growerProducts.length || 1).getValues();
    var packerRates = baseRatesSheet.getRange(3, 8, packerFruitTypes.length || 1, packerProducts.length || 1).getValues();

    var currencyData = discountsSheet.getRange('P3:Q' + discountsSheet.getLastRow()).getValues();
    var currencies = [];
    var rates = [];

    currencyData.forEach(row => {
      if (row[0] && !isNaN(parseFloat(row[1])) && isFinite(row[1])) {
        currencies.push(String(row[0]).trim());
        rates.push(parseFloat(row[1]));
      }
    });

    var regionData = discountsSheet.getRange('I3:K' + discountsSheet.getLastRow()).getValues();
    var regions = [];
    var growerRegionDiscounts = [];
    var packerRegionDiscounts = [];

    regionData.forEach(row => {
      var region = String(row[0] || '').trim();
      var growerDiscount = !isNaN(parseFloat(row[1])) && isFinite(row[1]) ? parseFloat(row[1]) : 0;
      var packerDiscount = !isNaN(parseFloat(row[2])) && isFinite(row[2]) ? parseFloat(row[2]) : 0;

      if (region) {
        regions.push(region);
        growerRegionDiscounts.push(growerDiscount);
        packerRegionDiscounts.push(packerDiscount);
      }
    });

    var paymentFrequencyData = discountsSheet.getRange('M3:N' + discountsSheet.getLastRow()).getValues();
    var paymentFrequencies = {};

    paymentFrequencyData.forEach(row => {
      var key = String(row[0] || '').trim();
      if (key && !isNaN(parseFloat(row[1])) && isFinite(row[1])) {
        paymentFrequencies[key] = parseFloat(row[1]);
      }
    });

    const lastRowWithDataA = discountsSheet.getRange("A:A").getValues().filter(String).length;
    const bulkDiscStartRow = 6;
    const numBulkRows = Math.max(1, lastRowWithDataA - bulkDiscStartRow + 1);
    var bulkDiscountRange = discountsSheet.getRange(bulkDiscStartRow, 1, numBulkRows, 7);
    var bulkDiscountData = bulkDiscountRange.getValues();

    var tieredBulkDiscounts = {};

    for (let i = 0; i < bulkDiscountData.length; i += 3) {
      const fruitTypeRow = bulkDiscountData[i];
      const percentageRow = (i + 1 < bulkDiscountData.length) ? bulkDiscountData[i+1] : null;

      if (!fruitTypeRow || !percentageRow) break;

      const fruitType = String(fruitTypeRow[0] || '').trim();
      if (!fruitType) continue;

      let tiers = [];

      for (let colIndex = 1; colIndex <= 6; colIndex++) {
        const thRaw = fruitTypeRow[colIndex];
        const perRaw = percentageRow[colIndex];

        if (thRaw !== null && thRaw !== '' && perRaw !== null && perRaw !== '') {
          try {
            const th = parseFloat(String(thRaw).replace(/[^0-9.-]+/g,""));
            const per = parseFloat(String(perRaw).replace(/[^0-9.-]+/g,""));

            if (!isNaN(th) && !isNaN(per)) {
              tiers.push({ threshold: th, percentage: per });
            }
          } catch (e) {}
        } else {
          break;
        }
      }

      if (tiers.length > 0) {
        tieredBulkDiscounts[fruitType] = tiers.sort((a, b) => a.threshold - b.threshold);
      }
    }

    const addOnHeaderRange = discountsSheet.getRange('V6:AD6');
    const addOnBaseProductsRaw = addOnHeaderRange.getValues().flat();
    const addOnBaseProducts = trimAndFilter(addOnBaseProductsRaw);

    const addOnDataStartRow = 7;
    const numAddOnCols = addOnBaseProducts.length + 1;
    const lastSheetRow = discountsSheet.getLastRow();
    const numAddOnRows = Math.max(1, lastSheetRow - addOnDataStartRow + 1);
    const addOnDataRange = discountsSheet.getRange(addOnDataStartRow, 22, numAddOnRows, numAddOnCols);
    const addOnData = addOnDataRange.getValues();

    var addOns = {};

    addOnData.forEach((row, index) => {
      const aoName = String(row[0] || '').trim();

      if (aoName) {
        addOns[aoName] = {};

        addOnBaseProducts.forEach((bp, pIdx) => {
          const raw = row[pIdx + 1];
          const mv = parseFloat(raw);

          if (bp && raw !== "" && !isNaN(mv)) {
            addOns[aoName][bp] = mv;
          }
        });

        if (Object.keys(addOns[aoName]).length === 0) {
          delete addOns[aoName];
        }
      }
    });

    var cameraRentalPriceNZD = 0;
    try {
      const raw = discountsSheet.getRange(CAMERA_PRICE_CELL).getValue();
      cameraRentalPriceNZD = parseFloat(raw);
      if (isNaN(cameraRentalPriceNZD)) {
        cameraRentalPriceNZD = 0;
      }
    } catch (e) {
      cameraRentalPriceNZD = 0;
    }

    var inflationRateDecimal = 0;
    try {
      const raw = discountsSheet.getRange(INFLATION_RATE_CELL).getValue();
      inflationRateDecimal = parseFloat(raw);
      if (isNaN(inflationRateDecimal)) {
        inflationRateDecimal = 0;
      }
    } catch(e) {
      inflationRateDecimal = 0;
    }

    var minimumPrices = {};
    const minPriceStartRow = MIN_PRICE_START_ROW;
    const lastRowWithDataS = discountsSheet.getLastRow();
    const numMinPriceRows = Math.max(0, lastRowWithDataS - minPriceStartRow + 1);

    if (numMinPriceRows > 0) {
      const minPriceRange = discountsSheet.getRange(minPriceStartRow, 19, numMinPriceRows, 2);
      const minPriceData = minPriceRange.getValues();

      minPriceData.forEach(row => {
        const productName = String(row[0] || '').trim();
        const minPriceTotal = parseFloat(row[1]);

        if (productName && !isNaN(minPriceTotal) && isFinite(minPriceTotal)) {
          minimumPrices[productName] = minPriceTotal;
        }
      });
    }
    // Read one-off add-on products from 'Base Rates per Fruit Type' Q2:R2 (names) and Q3:R3 (prices), all USD
    var oneOffAddOnProducts = [];
    try {
      var oneOffNamesRow = baseRatesSheet.getRange('Q2:R2').getValues()[0] || [];
      var oneOffPricesRow = baseRatesSheet.getRange('Q3:R3').getValues()[0] || [];
      for (var i = 0; i < oneOffNamesRow.length; i++) {
        var nm = String(oneOffNamesRow[i] || '').trim();
        var prRaw = oneOffPricesRow[i];
        var pr = parseFloat(String(prRaw).toString().replace(/[^0-9.-]+/g, ''));
        if (nm) {
          oneOffAddOnProducts.push({ name: nm, price: isNaN(pr) ? 0 : pr });
        }
      }
    } catch (e) {
      Logger.log('One-off add-on read error: ' + e);
    }
    
    // Read rental add-on products from 'Base Rates per Fruit Type' T2:U2 (names) and T3:U3 (prices), USD per year
    var rentalAddOnProducts = [];
    try {
      var rentalNamesRow = baseRatesSheet.getRange('T2:U2').getValues()[0] || [];
      var rentalPricesRow = baseRatesSheet.getRange('T3:U3').getValues()[0] || [];
      for (var j = 0; j < rentalNamesRow.length; j++) {
        var rnm = String(rentalNamesRow[j] || '').trim();
        var rprRaw = rentalPricesRow[j];
        var rpr = parseFloat(String(rprRaw).toString().replace(/[^0-9.-]+/g, ''));
        if (rnm) {
          rentalAddOnProducts.push({ name: rnm, price: isNaN(rpr) ? 0 : rpr });
        }
      }
    } catch (e) {
      Logger.log('Rental add-on read error: ' + e);
    }
    return { growerTypes: growerFruitTypes, packerTypes: packerFruitTypes, growerProducts: growerProducts, packerProducts: packerProducts, growerRates: growerRates, packerRates: packerRates, currencies: currencies, currencyRates: rates, regions: regions, growerRegionDiscounts: growerRegionDiscounts, packerRegionDiscounts: packerRegionDiscounts, paymentFrequencies: paymentFrequencies, tieredBulkDiscounts: tieredBulkDiscounts, addOns: addOns, cameraRentalPriceNZD: cameraRentalPriceNZD, inflationRateDecimal: inflationRateDecimal, minimumPrices: minimumPrices, oneOffAddOnProducts: oneOffAddOnProducts, rentalAddOnProducts: rentalAddOnProducts };
  } catch (e) { 
    Logger.log(`GetData Error: ${e}\n${e.stack}`); 
    return { error: `Data fetch failed: ${e.message}` }; 
  }
}

// --- Enhanced Calculation Function with Fixed Final Price Logic ---
function calculatePrice(formData) {
  try {
    var data = getData();
    if (data && data.error) { 
      Logger.log(data.error); 
      return { error: data.error }; 
    }
    if (!data || !data.minimumPrices || data.cameraRentalPriceNZD === undefined || data.inflationRateDecimal === undefined) { 
      return { error: "Failed to load critical data. Check logs." }; 
    }

    var customerType = formData.customerType; 
    var selectedFruits = formData.selectedFruits; 
    var region = formData.region; 
    var currency = formData.currency; 
    var paymentFrequencyKey = formData.paymentFrequency; 
    var includeCamera = formData.includeCamera; 
    var cameraCount = formData.cameraCount; 
    var discretionaryDiscountInputPercent = parseFloat(formData.discretionaryDiscount) || 0; 
    var discountFirstYearOnly = !!formData.discountFirstYearOnly;
    var contractYears = parseInt(formData.contractYears) || 1; 
    var companyName = formData.companyName || ""; 
    var companyAddress = formData.companyAddress || "";
    var salesContact = formData.salesContact || ""; 
    
    var fruitTypes = customerType === 'Grower' ? data.growerTypes : data.packerTypes; 
    var productsList = customerType === 'Grower' ? data.growerProducts : data.packerProducts; 
    var ratesTable = customerType === 'Grower' ? data.growerRates : data.packerRates; 
    
    var regionDiscountDecimal = 0; 
    var regionIndex = data.regions.indexOf(region); 
    if (regionIndex !== -1) { 
      regionDiscountDecimal = customerType === 'Grower' ? data.growerRegionDiscounts[regionIndex] : data.packerRegionDiscounts[regionIndex]; 
    } 
    
    var paymentFrequencyDiscountDecimal = data.paymentFrequencies[paymentFrequencyKey] || 0; 
    
    // --- STEP 1: GET CURRENCY RATE FIRST ---
    var currencyRate = 1.0; 
    var currencyIndex = data.currencies.indexOf(currency); 
    if (currencyIndex !== -1) { 
      currencyRate = data.currencyRates[currencyIndex]; 
    } 
    
    Logger.log(`=== CURRENCY CONVERSION (Applied First) ===`);
    Logger.log(`Currency: ${currency}, Rate: ${currencyRate}`);
    
    var sheetInflationRate = data.inflationRateDecimal; 
    var minimumTotalPrices = data.minimumPrices; 
    let discretionaryDiscountCalcDecimal = discretionaryDiscountInputPercent / 100;

    var baseTotal = 0; 
    var grandTotalTonnage = 0; 
    var productTonnages = {}; 
    var productBaseCosts = {}; 
    var productFinalPrices = {}; 
    var fruitGroupedData = {}; 
    var totalBulkDiscountCalculated = 0; 
    var totalMinPriceAdjustments = 0; 
    var planString = "";
    
    // Store selected add-ons by fruit for later calculation
    var selectedAddOnsByFruit = {};

    // --- STEP 2: Calculate Base Costs per Product & Apply Currency Conversion Immediately ---
    for (var fruitTypeKey in selectedFruits) {
      var fruitInfo = selectedFruits[fruitTypeKey];
      var fruitType = String(fruitTypeKey || '').trim();
      var productsData = fruitInfo.products || {};
      var selectedAddOns = fruitInfo.addOns || [];
      var fruitIndex = fruitTypes.indexOf(fruitType);

      if (fruitIndex === -1 || Object.keys(productsData).length === 0) continue;

      productBaseCosts[fruitType] = {};
      productFinalPrices[fruitType] = {};
      productTonnages[fruitType] = {};
      selectedAddOnsByFruit[fruitType] = selectedAddOns;

      let currentFruitTotalTonnage = 0;
      let currentFruitBasePrice = 0;
      let productDetailsForPlan = [];

      for (var productName in productsData) {
        let prodTon = parseFloat(productsData[productName]) || 0;
        if (prodTon <= 0) continue;

        productTonnages[fruitType][productName] = prodTon;
        grandTotalTonnage += prodTon;
        currentFruitTotalTonnage += prodTon;

        let productIndex = productsList.indexOf(productName);
        let productBaseCostNZD = 0;

        if (productIndex !== -1) {
          let rate = (ratesTable[fruitIndex]?.[productIndex] !== undefined) ? ratesTable[fruitIndex][productIndex] : null;
          if (rate !== null && !isNaN(parseFloat(rate)) && isFinite(rate)) {
            productBaseCostNZD = parseFloat(rate) * prodTon;
          }
        } else {
          // Fallback: product may belong to the other catalogue (Packer vs Grower)
          const altProductsList = customerType === 'Grower' ? (data.packerProducts || []) : (data.growerProducts || []);
          const altRatesTable = customerType === 'Grower' ? (data.packerRates || []) : (data.growerRates || []);
          const altTypes = customerType === 'Grower' ? (data.packerTypes || []) : (data.growerTypes || []);
          const altFruitIndex = altTypes.indexOf(fruitType);
          if (altFruitIndex !== -1) {
            const altIdx = altProductsList.indexOf(productName);
            if (altIdx !== -1) {
              const altRate = (altRatesTable[altFruitIndex]?.[altIdx] !== undefined) ? altRatesTable[altFruitIndex][altIdx] : null;
              if (altRate !== null && !isNaN(parseFloat(altRate)) && isFinite(altRate)) {
                productBaseCostNZD = parseFloat(altRate) * prodTon;
              }
            }
          }
        }

        // Enforce product minimum price rule (NZD): cost = max(baseRate*Tonnage, product minimum)
        const minNZDRaw = data.minimumPrices ? data.minimumPrices[productName] : null;
        const minNZD = (minNZDRaw !== undefined && minNZDRaw !== null && !isNaN(parseFloat(minNZDRaw))) ? parseFloat(minNZDRaw) : 0;
        if (minNZD > 0) {
          productBaseCostNZD = Math.max(productBaseCostNZD, minNZD);
        }

        // --- APPLY CURRENCY CONVERSION TO BASE COST IMMEDIATELY ---
        let productBaseCostConverted = productBaseCostNZD * currencyRate;
        Logger.log(`Product ${productName}: Base cost NZD=${productBaseCostNZD}, Converted=${productBaseCostConverted} ${currency}`);

        productBaseCosts[fruitType][productName] = productBaseCostConverted;
        currentFruitBasePrice += productBaseCostConverted;
        productDetailsForPlan.push(`${productName} (${prodTon.toFixed(2)} t)`);
      }

      baseTotal += currentFruitBasePrice;

      if (productDetailsForPlan.length > 0) {
        planString += `${fruitType}: ${productDetailsForPlan.join(', ')}\n`;
      }

      let pricePerTonneForFruit = (currentFruitTotalTonnage > 0) ? (currentFruitBasePrice / currentFruitTotalTonnage) : 0;
      let fruitBulkDiscountAmount = 0;
      let applicableTiers = data.tieredBulkDiscounts[fruitType];

      if (applicableTiers && applicableTiers.length > 0 && pricePerTonneForFruit > 0) {
        for (let i = 0; i < applicableTiers.length; i++) {
          let tier = applicableTiers[i];
          let thresh = tier.threshold;
          let perc = tier.percentage;
          let nextThresh = (i + 1 < applicableTiers.length) ? applicableTiers[i+1].threshold : Infinity;

          let tonnageInBracket = Math.max(0, Math.min(currentFruitTotalTonnage, nextThresh) - thresh);
          if (tonnageInBracket > 0) {
            fruitBulkDiscountAmount += tonnageInBracket * pricePerTonneForFruit * perc;
          }

          if (currentFruitTotalTonnage <= nextThresh) break;
        }
        totalBulkDiscountCalculated += fruitBulkDiscountAmount;
      }

      fruitGroupedData[fruitType] = { totalTonnage: currentFruitTotalTonnage, basePrice: currentFruitBasePrice, bulkDiscount: fruitBulkDiscountAmount };
    }

    // --- STEP 3: Apply NEW CASCADING DISCOUNT LOGIC First to get final product prices ---
    Logger.log(`=== NEW CASCADING DISCOUNT LOGIC (${currency}) ===`);
    
    let finalProductPricesConverted = {};
    let totalAppliedBulkDiscount = 0;
    let totalAppliedRegionDiscount = 0;
    let totalAppliedPaymentDiscount = 0;
    let totalAppliedDiscretionaryDiscount = 0;

    // Calculate Camera rental (convert to selected currency)
    var cameraRentalCost = 0;
    if (customerType === 'Packer' && includeCamera && cameraCount > 0) {
      cameraRentalCost = cameraCount * data.cameraRentalPriceNZD * currencyRate;
    }
    var roundedCameraRental = roundUpToNearestHundred(cameraRentalCost);

    for (let fruitType in productBaseCosts) {
      let fruitData = fruitGroupedData[fruitType];
      let fruitBase = fruitData.basePrice || 0;
      let fruitTotalBulkDiscount = fruitData.bulkDiscount || 0;
      
      finalProductPricesConverted[fruitType] = {};

      for (let productName in productBaseCosts[fruitType]) {
        if (productName.startsWith('_')) continue;

        let baseProductPrice = productBaseCosts[fruitType][productName]; // Already converted
        let productPortion = (fruitBase > 0) ? (baseProductPrice / fruitBase) : 0;
        
        Logger.log(`=== Product ${productName} in ${fruitType} (Currency: ${currency}) ===`);
        Logger.log(`Base price (converted): ${baseProductPrice} ${currency}`);
        
        // Calculate potential discount amounts (exact values, no rounding yet)
        let potentialBulkDiscount = fruitTotalBulkDiscount * productPortion;
        let potentialRegionDiscount = baseProductPrice * regionDiscountDecimal;
        let potentialPaymentDiscount = baseProductPrice * paymentFrequencyDiscountDecimal;
        let potentialDiscretionaryDiscount = baseProductPrice * discretionaryDiscountCalcDecimal;
        
        // --- APPLY CURRENCY CONVERSION TO MINIMUM PRICE ---
        let minTotalPriceValue = minimumTotalPrices[productName];
        let hasMinimumPrice = minTotalPriceValue !== undefined && minTotalPriceValue !== null && !isNaN(parseFloat(minTotalPriceValue));
        let minimumPriceConverted = hasMinimumPrice ? parseFloat(minTotalPriceValue) * currencyRate : 0;
        
        Logger.log(`Minimum price for ${productName}: NZD=${minTotalPriceValue}, Converted=${minimumPriceConverted} ${currency}`);
        
        let finalProductPrice = 0;
        let appliedBulkDiscount = 0;
        let appliedRegionDiscount = 0;
        let appliedPaymentDiscount = 0;
        let appliedDiscretionaryDiscount = 0;

        // NEW CASCADING LOGIC - Test with rounded camera (add-ons calculated later with rounded products)
        if (baseProductPrice >= minimumPriceConverted) {
          Logger.log(`Product ${productName}: Base price >= minimum, applying cascading logic`);
          
          // Try all discounts first - test with rounded camera only (add-ons calculated later)
          let testPrice = baseProductPrice + roundedCameraRental - potentialBulkDiscount - potentialRegionDiscount - potentialPaymentDiscount - potentialDiscretionaryDiscount;
          if (testPrice >= minimumPriceConverted) {
            // All discounts can be applied
            finalProductPrice = baseProductPrice;
            appliedBulkDiscount = potentialBulkDiscount;
            appliedRegionDiscount = potentialRegionDiscount;
            appliedPaymentDiscount = potentialPaymentDiscount;
            appliedDiscretionaryDiscount = potentialDiscretionaryDiscount;
            Logger.log(`All discounts applied for ${productName}`);
          } else {
            // Try without bulk discount
            testPrice = baseProductPrice + roundedCameraRental - potentialRegionDiscount - potentialPaymentDiscount - potentialDiscretionaryDiscount;
            if (testPrice >= minimumPriceConverted) {
              finalProductPrice = baseProductPrice;
              appliedRegionDiscount = potentialRegionDiscount;
              appliedPaymentDiscount = potentialPaymentDiscount;
              appliedDiscretionaryDiscount = potentialDiscretionaryDiscount;
              Logger.log(`Applied discounts except bulk for ${productName}`);
            } else {
              // Try without region discount
              testPrice = baseProductPrice + roundedCameraRental - potentialPaymentDiscount - potentialDiscretionaryDiscount;
              if (testPrice >= minimumPriceConverted) {
                finalProductPrice = baseProductPrice;
                appliedPaymentDiscount = potentialPaymentDiscount;
                appliedDiscretionaryDiscount = potentialDiscretionaryDiscount;
                Logger.log(`Applied payment and discretionary discounts only for ${productName}`);
              } else {
                // Try with only payment discount
                testPrice = baseProductPrice + roundedCameraRental - potentialPaymentDiscount;
                if (testPrice >= minimumPriceConverted) {
                  finalProductPrice = baseProductPrice;
                  appliedPaymentDiscount = potentialPaymentDiscount;
                  Logger.log(`Applied payment discount only for ${productName}`);
                } else {
                  // Use minimum price
                  finalProductPrice = minimumPriceConverted;
                  totalMinPriceAdjustments += (minimumPriceConverted - baseProductPrice);
                  Logger.log(`Using minimum price for ${productName}`);
                }
              }
            }
          }
        } else {
          // Product < Minimum price, return Minimum Price
          finalProductPrice = minimumPriceConverted;
          totalMinPriceAdjustments += (minimumPriceConverted - baseProductPrice);
          Logger.log(`Product ${productName}: Base price < minimum, using minimum price`);
        }

        Logger.log(`Final price for ${productName}: ${finalProductPrice} ${currency}`);
        
        finalProductPricesConverted[fruitType][productName] = finalProductPrice;
        
        // Accumulate applied discounts
        totalAppliedBulkDiscount += appliedBulkDiscount;
        totalAppliedRegionDiscount += appliedRegionDiscount;
        totalAppliedPaymentDiscount += appliedPaymentDiscount;
        totalAppliedDiscretionaryDiscount += appliedDiscretionaryDiscount;
      }
    }

    // --- STEP 4: Round up individual product prices and calculate rounded products total ---
    Logger.log(`=== ROUNDING INDIVIDUAL PRODUCTS (${currency}) ===`);
    
    let roundedProductsTotal = 0;
    
    for (let fruitType in finalProductPricesConverted) {
      for (let productName in finalProductPricesConverted[fruitType]) {
        let exactPrice = finalProductPricesConverted[fruitType][productName];
        let roundedPrice = roundUpToNearestHundred(exactPrice);
        finalProductPricesConverted[fruitType][productName] = roundedPrice;
        roundedProductsTotal += roundedPrice;
        Logger.log(`Product ${productName}: ${exactPrice} -> ${roundedPrice} (rounded up)`);
      }
    }

    // --- STEP 5: Calculate Add-ons based on ROUNDED product costs (MOVED AFTER ROUNDING) ---
    var addOnCosts = {};
    var exactAddOnCosts = {}; // Keep track of exact values
    var totalAddOnCost = 0;

    Logger.log(`=== ADD-ON CALCULATION START (${currency}) - Using ROUNDED product prices ===`);

    // Exclude one-off add-on products from recurring add-on calculations
    const oneOffNamesSet = new Set(((data && data.oneOffAddOnProducts) || []).map(p => String(p.name || '').trim()));

    for (let fruitType in selectedAddOnsByFruit) {
      let selectedAddOns = selectedAddOnsByFruit[fruitType] || [];
      let productsData = selectedFruits[fruitType].products || {};
      
      Logger.log(`Processing add-ons for ${fruitType}: ${JSON.stringify(selectedAddOns)}`);
      
      selectedAddOns.forEach(aoName => {
        const nameTrim = String(aoName || '').trim();
        if (oneOffNamesSet.has(nameTrim)) {
          Logger.log(`Skipping one-off add-on in recurring calc: ${nameTrim}`);
          return; // skip one-off items from recurring add-ons
        }
        const details = data.addOns[nameTrim];
        Logger.log(`Processing add-on: ${nameTrim}, details: ${JSON.stringify(details)}`);
        
        if (details) {
          let applicableRoundedProductsTotal = 0;
          
          // Calculate total of ROUNDED product prices for applicable products
          for (const pName in productsData) {
            Logger.log(`Checking product ${pName} for add-on ${nameTrim}`);
            if (details.hasOwnProperty(pName)) {
              Logger.log(`Product ${pName} is applicable for add-on ${nameTrim}`);
              if (finalProductPricesConverted[fruitType] && finalProductPricesConverted[fruitType][pName]) {
                const roundedProductPrice = finalProductPricesConverted[fruitType][pName]; // Already rounded
                applicableRoundedProductsTotal += roundedProductPrice;
                Logger.log(`Add-on ${nameTrim} - Product ${pName}: Rounded price ${roundedProductPrice} ${currency}, Running total: ${applicableRoundedProductsTotal} ${currency}`);
              }
            }
          }

          Logger.log(`Total applicable rounded products value for ${nameTrim}: ${applicableRoundedProductsTotal} ${currency}`);

          if (applicableRoundedProductsTotal > 0) {
            let markup = 0;
            for (const bp in details) {
              if (productsData.hasOwnProperty(bp)) {
                markup = details[bp];
                Logger.log(`Found markup for add-on ${nameTrim}: ${markup}`);
                break;
              }
            }

            if (markup !== 0) {
              // Calculate add-on cost as exact percentage markup of ROUNDED product prices
              const addonCostExact = applicableRoundedProductsTotal * markup;
              
              Logger.log(`Add-on ${nameTrim} calculation (${currency}): Rounded Base=${applicableRoundedProductsTotal}, Markup=${markup} (${markup*100}%), Result=${addonCostExact}`);
              
              // Store exact values for internal calculations
              exactAddOnCosts[nameTrim] = (exactAddOnCosts[nameTrim] || 0) + addonCostExact;
            }
          }
        }
      });
    }

    // Round individual add-ons and calculate total from rounded values
    for (let aoName in exactAddOnCosts) {
      addOnCosts[aoName] = roundUpToNearestHundred(exactAddOnCosts[aoName]);
      totalAddOnCost += addOnCosts[aoName];
    }

    Logger.log("=== ADD-ON CALCULATION END ===");
    Logger.log(`Final add-on costs (${currency}): ${JSON.stringify(addOnCosts)}`);
    Logger.log(`Total add-on cost (${currency}): ${totalAddOnCost}`);

    // --- STEP 6: Round up the discount amounts and calculate final cost ---
    var roundedBulkDiscount = roundUpToNearestHundred(totalAppliedBulkDiscount);
    var roundedRegionDiscount = roundUpToNearestHundred(totalAppliedRegionDiscount);
    var roundedPaymentDiscount = roundUpToNearestHundred(totalAppliedPaymentDiscount);
    var roundedDiscretionaryDiscount = roundUpToNearestHundred(totalAppliedDiscretionaryDiscount);
    
    Logger.log(`=== FINAL CALCULATION WITH ROUNDED VALUES (${currency}) ===`);
    Logger.log(`Products total (from rounded individual products): ${roundedProductsTotal}`);
    Logger.log(`Add-ons: ${totalAddOnCost} (calculated from rounded products)`);
    Logger.log(`Camera: ${cameraRentalCost} -> ${roundedCameraRental} (rounded up)`);
    Logger.log(`Bulk discount: ${totalAppliedBulkDiscount} -> ${roundedBulkDiscount} (rounded up)`);
    Logger.log(`Region discount: ${totalAppliedRegionDiscount} -> ${roundedRegionDiscount} (rounded up)`);
    Logger.log(`Payment discount: ${totalAppliedPaymentDiscount} -> ${roundedPaymentDiscount} (rounded up)`);
    Logger.log(`Discretionary discount: ${totalAppliedDiscretionaryDiscount} -> ${roundedDiscretionaryDiscount} (rounded up)`);
    
    // --- Calculate Final Year 1 Cost using ROUNDED VALUES (base, before one-off add-ons) ---
    var baseYear1Cost = roundedProductsTotal + totalAddOnCost + roundedCameraRental - roundedBulkDiscount - roundedRegionDiscount - roundedPaymentDiscount - roundedDiscretionaryDiscount;
    
    // One-off add-on products from Base Rates (USD), added last to Year 1 only with no discounts
    var oneOffAddOnTotalUSD = 0;
    var oneOffAddOnBreakdown = {};
    var oneOffAddOnQuantities = {};
    try {
      var oneOffSelections = Array.isArray(formData.oneOffAddOns) ? formData.oneOffAddOns : [];
      var oneOffCatalog = (data && data.oneOffAddOnProducts) || [];
      oneOffSelections.forEach(function(sel) {
        var nm = String(sel.name || '').trim();
        var qty = parseInt(sel.qty, 10) || 0;
        if (!nm || qty <= 0) return;
        var catalogItem = oneOffCatalog.find(function(it){ return String(it.name || '').trim() === nm; });
        var unit = catalogItem && typeof catalogItem.price !== 'undefined' ? (parseFloat(catalogItem.price) || 0) : 0;
        var unitRounded = Math.ceil(unit);
        var subtotalRounded = unitRounded * qty;
        if (subtotalRounded > 0) {
          oneOffAddOnTotalUSD += subtotalRounded;
          oneOffAddOnBreakdown[nm] = (oneOffAddOnBreakdown[nm] || 0) + subtotalRounded;
          oneOffAddOnQuantities[nm] = (oneOffAddOnQuantities[nm] || 0) + qty;
        }
      });
    } catch (e) {
      Logger.log('One-off add-on calc error: ' + e);
    }

    // Rental add-on products (USD per year), added last and repeated every year, no discounts
    var rentalAddOnTotalUSD = 0;
    var rentalAddOnBreakdown = {};
    var rentalAddOnQuantities = {};
    try {
      var rentalSelections = Array.isArray(formData.rentalAddOns) ? formData.rentalAddOns : [];
      var rentalCatalog = (data && data.rentalAddOnProducts) || [];
      rentalSelections.forEach(function(sel) {
        var nm = String(sel.name || '').trim();
        var qty = parseInt(sel.qty, 10) || 0;
        if (!nm || qty <= 0) return;
        var catalogItem = rentalCatalog.find(function(it){ return String(it.name || '').trim() === nm; });
        var unit = catalogItem && typeof catalogItem.price !== 'undefined' ? (parseFloat(catalogItem.price) || 0) : 0;
        var unitRounded = Math.ceil(unit);
        var subtotalRounded = unitRounded * qty;
        if (subtotalRounded > 0) {
          rentalAddOnTotalUSD += subtotalRounded;
          rentalAddOnBreakdown[nm] = (rentalAddOnBreakdown[nm] || 0) + subtotalRounded;
          rentalAddOnQuantities[nm] = (rentalAddOnQuantities[nm] || 0) + qty;
        }
      });
    } catch (e) {
      Logger.log('Rental add-on calc error: ' + e);
    }

    var finalYear1Cost = baseYear1Cost + rentalAddOnTotalUSD + oneOffAddOnTotalUSD;
    Logger.log(`Final Calculation (${currency}): Base(${baseYear1Cost}) + Rental(${rentalAddOnTotalUSD}) + One-off(${oneOffAddOnTotalUSD}) = ${finalYear1Cost}`);

    // Standard annual cost for years after Year 1 (exclude one-off add-ons; add back discretionary if first-year only)
    var standardAnnualCost = baseYear1Cost + rentalAddOnTotalUSD + (discountFirstYearOnly ? roundedDiscretionaryDiscount : 0);

    // --- STEP 7: Calculate Multi-Year Values (do not multiply one-off add-ons; rentals recur each year) ---
    var totalContractValue = finalYear1Cost + (contractYears > 1 ? standardAnnualCost * (contractYears - 1) : 0);
    var inflationSavings = 0;
    
    if (contractYears > 1) { 
      let rateForProjection = sheetInflationRate; 
      if (rateForProjection > 0) { 
        const r_proj = 1 + rateForProjection; 
        let projectedValueWithInflation;
        if (discountFirstYearOnly) {
          const firstTerm = baseYear1Cost + rentalAddOnTotalUSD;
          const recurringTerm = baseYear1Cost + roundedDiscretionaryDiscount + rentalAddOnTotalUSD;
          const sumExcludingFirst = (Math.pow(r_proj, contractYears) - r_proj) / (r_proj - 1); // r + r^2 + ... + r^(N-1)
          projectedValueWithInflation = firstTerm + recurringTerm * sumExcludingFirst + oneOffAddOnTotalUSD;
        } else {
          projectedValueWithInflation = (baseYear1Cost + rentalAddOnTotalUSD) * (1 - Math.pow(r_proj, contractYears)) / (1 - r_proj) + oneOffAddOnTotalUSD; 
        }
        inflationSavings = Math.max(0, projectedValueWithInflation - totalContractValue); 
      } 
    }

    // Prepare products array for frontend
    let productsArray = [];
    for (let fruitType in finalProductPricesConverted) {
      for (let productName in finalProductPricesConverted[fruitType]) {
        productsArray.push({
          name: productName,
          tonnage: productTonnages[fruitType] ? (productTonnages[fruitType][productName] || 0) : 0,
          price: finalProductPricesConverted[fruitType][productName],
          fruit: fruitType,
          minPriceAdjustment: 0 // Could be calculated if needed
        });
      }
    }

    // Prepare discounts array for frontend (negative amounts for discounts, positive for add-ons)
    let discountsArray = [];

    if (roundedBulkDiscount > 0) {
      discountsArray.push({ name: "Bulk Discount", percentage: 0, amount: -roundedBulkDiscount });
    }
    if (roundedRegionDiscount > 0) {
      discountsArray.push({ name: "Region Discount", percentage: regionDiscountDecimal, amount: -roundedRegionDiscount });
    }
    if (roundedPaymentDiscount > 0) {
      discountsArray.push({ name: "Payment Frequency Discount", percentage: paymentFrequencyDiscountDecimal, amount: -roundedPaymentDiscount });
    }
    if (roundedDiscretionaryDiscount > 0) {
      discountsArray.push({ name: "Discretionary Discount", percentage: discretionaryDiscountCalcDecimal, amount: -roundedDiscretionaryDiscount });
    }
    if (totalMinPriceAdjustments !== 0) {
      discountsArray.push({ name: "Minimum Price Adjustments", percentage: 0, amount: totalMinPriceAdjustments });
    }
    // Add-ons are positive amounts
    for (let aoName in addOnCosts) {
      discountsArray.push({ name: aoName, percentage: 0, amount: addOnCosts[aoName] });
    }
    // Include one-off add-on products (USD) as positive amounts
    for (let onm in oneOffAddOnBreakdown) {
      discountsArray.push({ name: onm, percentage: 0, amount: oneOffAddOnBreakdown[onm] });
    }

    // Determine if discretionary discount requires approval (>20%)
    let requiresApproval = discretionaryDiscountCalcDecimal > 0.20;

    return { 
      success: true, 
      customerType, 
      productFinalPrices: finalProductPricesConverted, 
      productTonnages: productTonnages, 
      baseTotal: baseTotal, 
      totalTonnage: grandTotalTonnage, 
      currency, 
      region, 
      paymentFrequencyKey, 
      contractYears, 
      regionDiscountPercentDecimal: regionDiscountDecimal, 
      paymentFrequencyDiscountPercentDecimal: paymentFrequencyDiscountDecimal, 
      discretionaryDiscountPercentDecimal: discretionaryDiscountCalcDecimal, 
      discretionaryFirstYearOnly: discountFirstYearOnly,
      bulkDiscountAmount: roundedBulkDiscount,
      regionDiscountAmount: roundedRegionDiscount, 
      paymentFrequencyDiscountAmount: roundedPaymentDiscount, 
      discretionaryDiscountAmount: roundedDiscretionaryDiscount, 
      cameraRental: roundedCameraRental, 
      addOnCosts: addOnCosts, 
      totalAddOnCost: totalAddOnCost, 
      oneOffAddOnCosts: oneOffAddOnBreakdown,
      oneOffAddOnQuantities: oneOffAddOnQuantities,
      oneOffAddOnTotalUSD: oneOffAddOnTotalUSD,
      rentalAddOnCosts: rentalAddOnBreakdown,
      rentalAddOnQuantities: rentalAddOnQuantities,
      rentalAddOnTotalUSD: rentalAddOnTotalUSD,
      finalYear1Cost: finalYear1Cost, 
      standardAnnualCost: standardAnnualCost,
      totalContractValue: totalContractValue, 
      inflationSavings: inflationSavings, 
      totalMinPriceAdjustments: totalMinPriceAdjustments, 
      finalTotal: finalYear1Cost, 
      planString: planString.trim(), 
      companyName: companyName, 
      companyAddress: companyAddress,
      salesContact: salesContact,
      products: productsArray,
      discounts: discountsArray,
      requiresApproval: requiresApproval,
      cameraCount: cameraCount
    };
  } catch (e) { 
    Logger.log(`Calc Error: ${e}\nStack: ${e.stack}`); 
    return { error: `Calculation failed: ${e.message}` }; 
  }
}

// --- Function to Create Docx Report ---
function createDocReport(resultData, templateId) {
  let tempFile = null; 
  try { 
    Logger.log("Starting createDocReport..."); 
    if (!templateId) { 
      throw new Error("Template ID missing."); 
    } 
    if (!resultData?.success) { 
      throw new Error("Invalid result data."); 
    } 
    
    const currency = resultData.currency || 'NZD'; 
    const formatCurrencyValue = (val) => (typeof val === 'number' && !isNaN(val)) ? `${currency} ${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : `${currency} 0.00`; 
    const formatPercentValue = (dec) => (typeof dec === 'number' && !isNaN(dec)) ? `${(dec * 100).toFixed(1)}%` : '0.0%'; 
    const formatAmount = (val) => (typeof val === 'number' && !isNaN(val) && val !== 0) ? `${val < 0 ? '-' : '+'} ${currency} ${Math.abs(val).toFixed(2)}` : 'N/A'; 
    const naIfEmpty = (val) => (val && String(val).trim() !== '') ? String(val).trim() : 'N/A'; 
    
    // Calculate rounded products total from individual rounded product prices
    let roundedProductsTotal = 0;
    if (resultData.productFinalPrices) {
      for (let fruitType in resultData.productFinalPrices) {
        for (let productName in resultData.productFinalPrices[fruitType]) {
          roundedProductsTotal += resultData.productFinalPrices[fruitType][productName];
        }
      }
    }
    
    // Create add-on plan string
    let addOnPlanString = 'N/A';
    if (resultData.addOnCosts && Object.keys(resultData.addOnCosts).length > 0) {
      const addOnNames = Object.keys(resultData.addOnCosts).filter(name => resultData.addOnCosts[name] > 0);
      if (addOnNames.length > 0) {
        addOnPlanString = addOnNames.join(', ');
      }
    }
    
    // Create detailed product breakdown by fruit type
    let productBreakdownString = 'N/A';
    if (resultData.productFinalPrices && resultData.productTonnages) {
      let breakdownLines = [];
      for (let fruitType in resultData.productFinalPrices) {
        let fruitProducts = [];
        let fruitTotal = 0;
        
        for (let productName in resultData.productFinalPrices[fruitType]) {
          const price = resultData.productFinalPrices[fruitType][productName];
          const tonnage = resultData.productTonnages[fruitType] ? (resultData.productTonnages[fruitType][productName] || 0) : 0;
          fruitTotal += price;
          fruitProducts.push(`  ${productName} (${tonnage.toFixed(1)}t): ${formatCurrencyValue(price)}`);
        }
        
        if (fruitProducts.length > 0) {
          breakdownLines.push(`${fruitType} - Total: ${formatCurrencyValue(fruitTotal)}`);
          breakdownLines.push(...fruitProducts);
          breakdownLines.push(''); // Empty line between fruit types
        }
      }
      
      if (breakdownLines.length > 0) {
        productBreakdownString = breakdownLines.join('\n');
      }
    }
    
    // Create detailed add-on breakdown by product
    let addOnBreakdownString = 'N/A';
    if (resultData.addOnCosts && Object.keys(resultData.addOnCosts).length > 0) {
      let addOnLines = [];
      // Add total first
      addOnLines.push(`Total Add-ons: ${formatCurrencyValue(resultData.totalAddOnCost || 0)}`);
      for (let addOnName in resultData.addOnCosts) {
        const cost = resultData.addOnCosts[addOnName];
        if (cost > 0) {
          addOnLines.push(`${addOnName}: ${formatCurrencyValue(cost)}`);
        }
      }
      if (addOnLines.length > 1) {
        addOnBreakdownString = addOnLines.join('\n');
      }
    }

    // Build detailed strings for one-off and annual add-ons (product (quantity): currency total)
    let oneOffAddOnDetailsString = 'N/A';
    if (resultData.oneOffAddOnCosts && resultData.oneOffAddOnQuantities) {
      const lines = Object.keys(resultData.oneOffAddOnCosts).map(name => {
        const qty = resultData.oneOffAddOnQuantities[name] || 0;
        const total = resultData.oneOffAddOnCosts[name] || 0;
        if (qty > 0 && total > 0) {
          const unitRounded = Math.ceil(total / qty);
          const totalRounded = unitRounded * qty;
          return `${name} (${qty}): ${formatCurrencyValue(totalRounded)}`;
        }
        return null;
      }).filter(Boolean);
      if (lines.length) oneOffAddOnDetailsString = lines.join('\n');
    }

    let annualAddOnDetailsString = 'N/A';
    if (resultData.rentalAddOnCosts && resultData.rentalAddOnQuantities) {
      const lines = Object.keys(resultData.rentalAddOnCosts).map(name => {
        const qty = resultData.rentalAddOnQuantities[name] || 0;
        const total = resultData.rentalAddOnCosts[name] || 0;
        if (qty > 0 && total > 0) {
          const unitRounded = Math.ceil(total / qty);
          const totalRounded = unitRounded * qty;
          return `${name} (${qty}): ${formatCurrencyValue(totalRounded)}`;
        }
        return null;
      }).filter(Boolean);
      if (lines.length) annualAddOnDetailsString = lines.join('\n');
    }
    
    // Calculate discount total: Products + Add-ons + Camera - Final Year 1 Cost
    const discountTotal = roundedProductsTotal + (resultData.totalAddOnCost || 0) + (resultData.cameraRental || 0) - (resultData.finalYear1Cost || 0);
    
    // Determine camera information - use blank strings instead of "N/A"
    const cameraCount = resultData.cameraCount || 0;
    const hasCameras = resultData.cameraRental > 0;
    const cameraCountText = hasCameras ? cameraCount.toString() : '';
    const cameraCostText = hasCameras ? formatCurrencyValue(resultData.cameraRental) : '';
    const cameraRentalLabel = hasCameras ? 'Camera Rental:' : '';
    
    // Annual discounts sum (region + payment + discretionary if not first-year-only)
    const annualDiscountSum = (resultData.regionDiscountAmount || 0) + (resultData.paymentFrequencyDiscountAmount || 0) + (!resultData.discretionaryFirstYearOnly ? (resultData.discretionaryDiscountAmount || 0) : 0);
    const discountsLabel = annualDiscountSum > 0 ? `Discounts: ${formatCurrencyValue(annualDiscountSum)}` : '';
    
    const placeholders = { 
      '{{CalculationDate}}': new Date().toLocaleDateString('en-NZ', { year: 'numeric', month: 'short', day: 'numeric'}), 
      '{{CustomerType}}': resultData.customerType || 'N/A', 
      '{{Company}}': naIfEmpty(resultData.companyName), 
      '{{CompanyAddress}}': naIfEmpty(resultData.companyAddress),
      '{{Sales}}': naIfEmpty(resultData.salesContact), 
      '{{Plan}}': naIfEmpty(resultData.planString), 
      '{{AddOnPlan}}': addOnPlanString,
      '{{ProductBreakdown}}': productBreakdownString,
      '{{DiscountTotal}}': formatCurrencyValue(discountTotal),
      '{{Discounts}}': discountsLabel,
      '{{Region}}': resultData.region || 'N/A', 
      '{{Currency}}': currency, 
      '{{ContractYears}}': resultData.contractYears || 1, 
      '{{PaymentFrequency}}': resultData.paymentFrequencyKey || 'N/A', 
      '{{BaseTotal}}': formatCurrencyValue(roundedProductsTotal), 
      '{{BulkDiscountAmount}}': formatAmount(-resultData.bulkDiscountAmount), 
      '{{RegionDiscountPercent}}': formatPercentValue(resultData.regionDiscountPercentDecimal), 
      '{{RegionDiscountAmount}}': formatAmount(-resultData.regionDiscountAmount), 
      '{{PaymentDiscountPercent}}': formatPercentValue(resultData.paymentFrequencyDiscountPercentDecimal), 
      '{{PaymentDiscountAmount}}': formatAmount(-resultData.paymentFrequencyDiscountAmount), 
      '{{DiscretionaryDiscountPercent}}': formatPercentValue(resultData.discretionaryDiscountPercentDecimal), 
      '{{DiscretionaryDiscountAmount}}': formatAmount(-resultData.discretionaryDiscountAmount), 
      '{{Y1discount}}': (resultData.discretionaryFirstYearOnly && resultData.discretionaryDiscountPercentDecimal > 0) ? `Year 1 Discount: ${formatPercentValue(resultData.discretionaryDiscountPercentDecimal)}` : '', 
      '{{MinPriceAdjustment}}': formatAmount(resultData.totalMinPriceAdjustments), 
      '{{AddonTotal}}': formatCurrencyValue(resultData.totalAddOnCost), 
      '{{CameraRental}}': formatAmount(resultData.cameraRental), 
      '{{Year1Cost}}': formatCurrencyValue(resultData.finalYear1Cost), 
      '{{StandardAnnualCost}}': formatCurrencyValue(resultData.standardAnnualCost), 
      '{{InflationSavings}}': resultData.contractYears > 1 && resultData.inflationSavings > 0 ? formatCurrencyValue(resultData.inflationSavings) : 'N/A', 
      '{{TotalContractValue}}': resultData.contractYears > 1 ? formatCurrencyValue(resultData.totalContractValue) : 'N/A', 
      '{{InflationStatus}}': resultData.contractYears > 1 ? 'Inflation Waived (Multi-Year Commitment)' : '1-Year Term', 
      '{{OneOffAddOnBreakdown}}': oneOffAddOnDetailsString,
      '{{AnnualAddOnBreakdown}}': annualAddOnDetailsString,
    };

    Logger.log("Accessing template file via DriveApp, ID: " + templateId); 
    const templateFile = DriveApp.getFileById(templateId); 
    if (!templateFile) { 
      throw new Error("Could not find template file."); 
    } 
    
    Logger.log("Template Name: '" + templateFile.getName() + "'"); 
    const tempFolderName = "Temporary Price Reports"; 
    var destFolder; 
    const folders = DriveApp.getFoldersByName(tempFolderName); 
    if (folders.hasNext()) { 
      destFolder = folders.next(); 
    } else { 
      destFolder = DriveApp.createFolder(tempFolderName); 
    } 
    
    Logger.log("Using destination folder: " + destFolder.getName()); 
    const docName = `Price Report - ${placeholders['{{Company}}'] !== 'N/A' ? placeholders['{{Company}}'] : 'Customer'} - ${placeholders['{{CalculationDate}}']}`; 
    tempFile = templateFile.makeCopy(docName, destFolder); 
    const copyId = tempFile.getId(); 
    Logger.log("Copy created: ID " + copyId); 
    
    const tempDoc = DocumentApp.openById(copyId); 
    const body = tempDoc.getBody(); 
    Logger.log("Replacing placeholders..."); 
    for (const key in placeholders) { 
      body.replaceText(key, placeholders[key] || ''); 
    } 
    tempDoc.saveAndClose(); 
    Logger.log("Doc saved.");

    Logger.log("Getting OAuth token..."); 
    const oauthToken = ScriptApp.getOAuthToken(); 
    const docxExportUrl = `https://docs.google.com/document/d/${copyId}/export?format=docx`; 
    Logger.log("Fetching Docx export URL..."); 
    const options = { headers: { 'Authorization': 'Bearer ' + oauthToken }, muteHttpExceptions: true }; 
    const response = UrlFetchApp.fetch(docxExportUrl, options); 
    const responseCode = response.getResponseCode(); 
    Logger.log("UrlFetch response code: " + responseCode); 
    if (responseCode !== 200) { 
      throw new Error(`Failed to export document (HTTP ${responseCode}): ${response.getContentText()}`); 
    } 
    
    const docxBlob = response.getBlob(); 
    Logger.log("Got blob. Encoding..."); 
    const base64Data = Utilities.base64Encode(docxBlob.getBytes()); 
    Logger.log("Encoding complete."); 
    Logger.log("Trashing temp file..."); 
    tempFile.setTrashed(true);

    const filename = `${docName}.docx`; 
    Logger.log("Returning data for: " + filename); 
    return { base64: base64Data, filename: filename, contentType: MimeType.MICROSOFT_WORD };
  } catch (e) { 
    Logger.log(`Error in createDocReport: ${e}`); 
    Logger.log(`Stack: ${e.stack}`); 
    if (tempFile) { 
      try { 
        Logger.log("Cleanup after error..."); 
        DriveApp.getFileById(tempFile.getId()).setTrashed(true); 
      } catch (err) {} 
    } 
    
    let errorMessage = e.message; 
    if (e.message.includes("makeCopy") || e.message.includes("getFileById") || e.message.includes("openById")) { 
      errorMessage = "Could not access/copy template. Check ID/Permissions."; 
    } else if (e.message.includes("UrlFetchApp")) { 
      errorMessage = "Could not export doc. Check script permissions."; 
    } 
    return { error: `Failed to create report: ${errorMessage}` }; 
  }
}

// --- Function to Create Google Doc Report ---
function createGoogleDocReport(resultData, templateId) {
  let tempFile = null; 
  try { 
    Logger.log("Starting createGoogleDocReport..."); 
    if (!templateId) { 
      throw new Error("Template ID missing."); 
    } 
    if (!resultData?.success) { 
      throw new Error("Invalid result data."); 
    } 
    
    const currency = resultData.currency || 'NZD'; 
    const formatCurrencyValue = (val) => (typeof val === 'number' && !isNaN(val)) ? `${currency} ${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : `${currency} 0.00`; 
    const formatPercentValue = (dec) => (typeof dec === 'number' && !isNaN(dec)) ? `${(dec * 100).toFixed(1)}%` : '0.0%'; 
    const formatAmount = (val) => (typeof val === 'number' && !isNaN(val) && val !== 0) ? `${val < 0 ? '-' : '+'} ${currency} ${Math.abs(val).toFixed(2)}` : 'N/A'; 
    const naIfEmpty = (val) => (val && String(val).trim() !== '') ? String(val).trim() : 'N/A'; 
    
    // Calculate rounded products total from individual rounded product prices
    let roundedProductsTotal = 0;
    if (resultData.productFinalPrices) {
      for (let fruitType in resultData.productFinalPrices) {
        for (let productName in resultData.productFinalPrices[fruitType]) {
          roundedProductsTotal += resultData.productFinalPrices[fruitType][productName];
        }
      }
    }
    
    // Create add-on plan string
    let addOnPlanString = 'N/A';
    if (resultData.addOnCosts && Object.keys(resultData.addOnCosts).length > 0) {
      const addOnNames = Object.keys(resultData.addOnCosts).filter(name => resultData.addOnCosts[name] > 0);
      if (addOnNames.length > 0) {
        addOnPlanString = addOnNames.join(', ');
      }
    }
    
    // Create detailed product breakdown by fruit type
    let productBreakdownString = 'N/A';
    if (resultData.productFinalPrices && resultData.productTonnages) {
      let breakdownLines = [];
      for (let fruitType in resultData.productFinalPrices) {
        let fruitProducts = [];
        let fruitTotal = 0;
        
        for (let productName in resultData.productFinalPrices[fruitType]) {
          const price = resultData.productFinalPrices[fruitType][productName];
          const tonnage = resultData.productTonnages[fruitType] ? (resultData.productTonnages[fruitType][productName] || 0) : 0;
          fruitTotal += price;
          fruitProducts.push(`  ${productName} (${tonnage.toFixed(1)}t): ${formatCurrencyValue(price)}`);
        }
        
        if (fruitProducts.length > 0) {
          breakdownLines.push(`${fruitType} - Total: ${formatCurrencyValue(fruitTotal)}`);
          breakdownLines.push(...fruitProducts);
          breakdownLines.push(''); // Empty line between fruit types
        }
      }
      
      if (breakdownLines.length > 0) {
        productBreakdownString = breakdownLines.join('\n');
      }
    }
    
    // Create detailed add-on breakdown by product
    let addOnBreakdownString = 'N/A';
    if (resultData.addOnCosts && Object.keys(resultData.addOnCosts).length > 0) {
      let addOnLines = [];
      // Add total first
      addOnLines.push(`Total Add-ons: ${formatCurrencyValue(resultData.totalAddOnCost || 0)}`);
      for (let addOnName in resultData.addOnCosts) {
        const cost = resultData.addOnCosts[addOnName];
        if (cost > 0) {
          addOnLines.push(`${addOnName}: ${formatCurrencyValue(cost)}`);
        }
      }
      if (addOnLines.length > 1) {
        addOnBreakdownString = addOnLines.join('\n');
      }
    }

    // Build detailed strings for one-off and annual add-ons (product (quantity): currency total)
    let oneOffAddOnDetailsString = 'N/A';
    if (Array.isArray(resultData.oneOffAddOnDetails) && resultData.oneOffAddOnDetails.length > 0) {
      const lines = resultData.oneOffAddOnDetails
        .filter(d => d && d.quantity > 0 && d.total > 0)
        .map(d => {
          const unitRounded = Math.ceil(d.total / d.quantity);
          const totalRounded = unitRounded * d.quantity;
          return `${d.name} (${d.quantity}): ${formatCurrencyValue(totalRounded)}`;
        });
      if (lines.length) oneOffAddOnDetailsString = lines.join('\n');
    }

    let annualAddOnDetailsString = 'N/A';
    if (Array.isArray(resultData.rentalAddOnDetails) && resultData.rentalAddOnDetails.length > 0) {
      const lines = resultData.rentalAddOnDetails
        .filter(d => d && d.quantity > 0 && d.total > 0)
        .map(d => {
          const unitRounded = Math.ceil(d.total / d.quantity);
          const totalRounded = unitRounded * d.quantity;
          return `${d.name} (${d.quantity}): ${formatCurrencyValue(totalRounded)}`;
        });
      if (lines.length) annualAddOnDetailsString = lines.join('\n');
    }
    
    // Calculate discount total: Products + Add-ons + Camera - Final Year 1 Cost
    const discountTotal = roundedProductsTotal + (resultData.totalAddOnCost || 0) + (resultData.cameraRental || 0) - (resultData.finalYear1Cost || 0);
    
    // Determine camera information - use blank strings instead of "N/A"
    const cameraCount = resultData.cameraCount || 0;
    const hasCameras = resultData.cameraRental > 0;
    const cameraCountText = hasCameras ? cameraCount.toString() : '';
    const cameraCostText = hasCameras ? formatCurrencyValue(resultData.cameraRental) : '';
    const cameraRentalLabel = hasCameras ? 'Camera Rental:' : '';
    
    // Annual discounts sum (region + payment + discretionary if not first-year-only)
    const annualDiscountSum = (resultData.regionDiscountAmount || 0) + (resultData.paymentFrequencyDiscountAmount || 0) + (!resultData.discretionaryFirstYearOnly ? (resultData.discretionaryDiscountAmount || 0) : 0);
    const discountsLabel = annualDiscountSum > 0 ? `Discounts: ${formatCurrencyValue(annualDiscountSum)}` : '';
    
    const placeholders = { 
      '{{CalculationDate}}': new Date().toLocaleDateString('en-NZ', { year: 'numeric', month: 'short', day: 'numeric'}), 
      '{{CustomerType}}': resultData.customerType || 'N/A', 
      '{{Company}}': naIfEmpty(resultData.companyName), 
      '{{CompanyAddress}}': naIfEmpty(resultData.companyAddress),
      '{{Sales}}': naIfEmpty(resultData.salesContact), 
      '{{Plan}}': naIfEmpty(resultData.planString), 
      '{{AddOnPlan}}': addOnPlanString,
      '{{ProductBreakdown}}': productBreakdownString,
      '{{DiscountTotal}}': formatCurrencyValue(discountTotal),
      '{{Discounts}}': discountsLabel,
      '{{Region}}': resultData.region || 'N/A', 
      '{{Currency}}': currency, 
      '{{ContractYears}}': resultData.contractYears || 1, 
      '{{PaymentFrequency}}': resultData.paymentFrequencyKey || 'N/A', 
      '{{BaseTotal}}': formatCurrencyValue(roundedProductsTotal), 
      '{{BulkDiscountAmount}}': formatAmount(-resultData.bulkDiscountAmount), 
      '{{RegionDiscountPercent}}': formatPercentValue(resultData.regionDiscountPercentDecimal), 
      '{{RegionDiscountAmount}}': formatAmount(-resultData.regionDiscountAmount), 
      '{{PaymentDiscountPercent}}': formatPercentValue(resultData.paymentFrequencyDiscountPercentDecimal), 
      '{{PaymentDiscountAmount}}': formatAmount(-resultData.paymentFrequencyDiscountAmount), 
      '{{DiscretionaryDiscountPercent}}': formatPercentValue(resultData.discretionaryDiscountPercentDecimal), 
      '{{DiscretionaryDiscountAmount}}': formatAmount(-resultData.discretionaryDiscountAmount), 
      '{{Y1discount}}': (resultData.discretionaryFirstYearOnly && resultData.discretionaryDiscountPercentDecimal > 0) ? `Year 1 Discount: ${formatPercentValue(resultData.discretionaryDiscountPercentDecimal)}` : '', 
      '{{MinPriceAdjustment}}': formatAmount(resultData.totalMinPriceAdjustments), 
      '{{AddonTotal}}': formatCurrencyValue(resultData.totalAddOnCost), 
      '{{CameraRental}}': formatAmount(resultData.cameraRental), 
      '{{Year1Cost}}': formatCurrencyValue(resultData.finalYear1Cost), 
      '{{StandardAnnualCost}}': formatCurrencyValue(resultData.standardAnnualCost), 
      '{{InflationSavings}}': resultData.contractYears > 1 && resultData.inflationSavings > 0 ? formatCurrencyValue(resultData.inflationSavings) : 'N/A', 
      '{{TotalContractValue}}': resultData.contractYears > 1 ? formatCurrencyValue(resultData.totalContractValue) : 'N/A', 
      '{{InflationStatus}}': resultData.contractYears > 1 ? 'Inflation Waived (Multi-Year Commitment)' : '1-Year Term', 
      '{{OneOffAddOnBreakdown}}': oneOffAddOnDetailsString,
      '{{AnnualAddOnBreakdown}}': annualAddOnDetailsString,
    };

    Logger.log("Accessing template file via DriveApp, ID: " + templateId); 
    const templateFile = DriveApp.getFileById(templateId); 
    if (!templateFile) { 
      throw new Error("Could not find template file."); 
    } 
    
    Logger.log("Template Name: '" + templateFile.getName() + "'"); 
    const tempFolderName = "Temporary Price Reports"; 
    var destFolder; 
    const folders = DriveApp.getFoldersByName(tempFolderName); 
    if (folders.hasNext()) { 
      destFolder = folders.next(); 
    } else { 
      destFolder = DriveApp.createFolder(tempFolderName); 
    } 
    
    Logger.log("Using destination folder: " + destFolder.getName()); 
    const docName = `Price Report - ${placeholders['{{Company}}'] !== 'N/A' ? placeholders['{{Company}}'] : 'Customer'} - ${placeholders['{{CalculationDate}}']}`; 
    tempFile = templateFile.makeCopy(docName, destFolder); 
    const copyId = tempFile.getId(); 
    Logger.log("Copy created: ID " + copyId); 
    
    const tempDoc = DocumentApp.openById(copyId); 
    const body = tempDoc.getBody(); 
    Logger.log("Replacing placeholders..."); 
    for (const key in placeholders) { 
      body.replaceText(key, placeholders[key] || ''); 
    } 
    tempDoc.saveAndClose(); 
    Logger.log("Google Doc saved.");

    // Get the URL of the created Google Doc
    const docUrl = tempFile.getUrl();
    
    const filename = docName; 
    Logger.log("Returning Google Doc URL for: " + filename); 
    return { url: docUrl, filename: filename };
  } catch (e) { 
    Logger.log(`Error in createGoogleDocReport: ${e}`); 
    Logger.log(`Stack: ${e.stack}`); 
    if (tempFile) { 
      try { 
        Logger.log("Cleanup after error..."); 
        DriveApp.getFileById(tempFile.getId()).setTrashed(true); 
      } catch (err) {} 
    } 
    
    let errorMessage = e.message; 
    if (e.message.includes("makeCopy") || e.message.includes("getFileById") || e.message.includes("openById")) { 
      errorMessage = "Could not access/copy template. Check ID/Permissions."; 
    }
    return { error: `Failed to create Google Doc report: ${errorMessage}` }; 
  }
}

// --- Debug Function ---
function displayResultsInSheet_DEBUG(result) {
  Logger.log("--- displayResultsInSheet_DEBUG ---");

  if (result.error) {
    Logger.log("Error: " + result.error);
    return false;
  }

  if (!result.success) {
    Logger.log("Success flag false.");
    return false;
  }

  Logger.log("Year 1 Cost: " + result.finalYear1Cost);
  Logger.log("TCV: " + result.totalContractValue);
  Logger.log("Savings: " + result.inflationSavings);
  Logger.log("Min Price Adjustments: " + result.totalMinPriceAdjustments);
  Logger.log("Product Breakdown: " + JSON.stringify(result.productFinalPrices));

  return true;
}