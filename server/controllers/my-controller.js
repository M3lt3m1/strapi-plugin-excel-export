"use strict";
const ExcelJS = require("exceljs");

module.exports = ({ strapi }) => ({

  // select the collections that are defined into "excel" configuration
  async getDropDownData() {
    let excel = strapi.config.get("excel");
    let dropDownValues = [];
    let array = Object.keys(excel?.config);

    // scan all the collections
    strapi?.db?.config?.models?.forEach((element) => {
      if (element?.kind == "collectionType") {
        array?.forEach((data) => {
          if (element?.uid?.startsWith(data)) {
            dropDownValues.push({
              label: element?.info?.displayName,
              value: element?.uid,
            });
          }
        });
      }
    });

    // Sort dropDownValues alphabetically by label in ascending order
    dropDownValues.sort((a, b) => a.label.localeCompare(b.label));
    return {
      data: dropDownValues,
    };
  },

  // extract the data from current collection
  async getTableData(ctx) {
    let excel = strapi.config.get("excel");
    let uid = ctx?.query?.uid;
    let limit = ctx?.query?.limit;
    let offset = ctx?.query?.offset;
    let query = await this.restructureObject(
      excel?.config[uid],
      uid,
      limit,
      offset
    );

    // query the data from the collection "uid"
    let response = await strapi.db.query(uid).findMany(query);

    // build the header
    let header = [
      ...excel?.config[uid]?.columns,                 // add collection field names
      ...Object.keys(excel?.config[uid]?.relation),   // add field names from relations
    ];

    let where = {};

    if (excel?.config[uid]?.locale == "true") {
      where = {
        locale: "en",
      };
    }

    // find how many occurrences has been found with current query
    let count = await strapi.db.query(uid).count(where);

    // format loaded data
    let tableData = await this.restructureData(response, excel?.config[uid]);

    // Sort dropDownValues alphabetically by label in ascending order

    return {
      data: tableData,
      count: count,
      columns: header,
    };
  },

  async downloadExcel(ctx) {
    try {
      let excel = strapi.config.get("excel");

      strapi.log.info(`export.downloadExcel: config=${JSON.stringify(excel, null, 2)}`);

      let uid = ctx?.query?.uid;

      strapi.log.info(`export.downloadExcel: UID=${uid}`);

      let query = await this.restructureObject(excel?.config[uid], uid);

      strapi.log.info(`export.downloadExcel: query=${JSON.stringify(query, null, 2)}`);

      let response = await strapi.db.query(uid).findMany(query);

      strapi.log.info(`export.downloadExcel: response=${JSON.stringify(response, null, 2).substring(0,500)}...`);

      let excelData = await this.restructureData(response, excel?.config[uid]);

      strapi.log.info(`export.downloadExcel: excelData=${JSON.stringify(excelData, null, 2).substring(0,500)}...`);

      // Create a new workbook and add a worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      // Extract column headers dynamically from the data
      let headers = [
        ...excel?.config[uid]?.columns,
        ...Object.keys(excel?.config[uid]?.relation),
      ];

      // Transform the original headers to the desired format
      let headerRestructure = [];
      headers?.forEach((element) => {
        const formattedHeader = element
          .split("_")
          .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
          .join(" ");
        headerRestructure.push(formattedHeader);
      });

      // Define dynamic column headers
      worksheet.columns = headers.map((header, index) => ({
        header: headerRestructure[index], // Use the formatted header
        key: header,
        width: 20,
      }));

      // Define the dropdown list options for the Gender column

      // Add data to the worksheet
      excelData?.forEach((row) => {
        // Excel will provide a dropdown with these values.
        worksheet.addRow(row);
      });

      // Enable text wrapping for all columns
      worksheet.columns.forEach((column) => {
        column.alignment = { wrapText: true };
      });

      // Freeze the first row
      worksheet.views = [
        { state: "frozen", xSplit: 0, ySplit: 1, topLeftCell: "A" },
      ];

      // Write the workbook to a file
      const buffer = await workbook.xlsx.writeBuffer();

      return buffer;
    } catch (error) {
      console.error("export.downloadExcelError writing buffer:", error);
    }
  },

  async downloadCSV(ctx) {
    try {
      let excel = strapi.config.get("excel");

      strapi.log.info(`export.downloadExcel: config=${JSON.stringify(excel, null, 2)}`);

      let uid = ctx?.query?.uid;

      strapi.log.info(`export.downloadCSV: UID=${uid}`);

      // make the query
      let query = await this.restructureObject(excel?.config[uid], uid);

      strapi.log.info(`export.downloadCSV: query=${JSON.stringify(query, null, 2)}`);

      // exec the query
      let response = await strapi.db.query(uid).findMany(query);

      strapi.log.info(`export.downloadCSV: response=${JSON.stringify(response, null, 2).substring(0,500)}...`);

      // format the data
      let excelData = await this.restructureData(response, excel?.config[uid]);

      strapi.log.info(`export.downloadCSV: excelData=${JSON.stringify(excelData, null, 2).substring(0,500)}...`);

      // Create a new workbook and add a worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      // Extract column headers dynamically from the data
      let headers = [
        ...excel?.config[uid]?.columns,
        ...Object.keys(excel?.config[uid]?.relation),
      ];

      // Transform the original headers to the desired format
      let headerRestructure = [];
      headers?.forEach((element) => {
        const formattedHeader = element
          .split("_")
          .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
          .join(" ");
        headerRestructure.push(formattedHeader);
      });

      // Define dynamic column headers
      worksheet.columns = headers.map((header, index) => ({
        header: headerRestructure[index], // Use the formatted header
        key: header,
        width: 20,
      }));

      // Define the dropdown list options for the Gender column

      // Add data to the worksheet
      excelData?.forEach((row) => {
        // Excel will provide a dropdown with these values.
        worksheet.addRow(row);
      });

      // Enable text wrapping for all columns
      worksheet.columns.forEach((column) => {
        column.alignment = { wrapText: true };
      });

      // Freeze the first row
      worksheet.views = [
        { state: "frozen", xSplit: 0, ySplit: 1, topLeftCell: "A" },
      ];

      // Write the workbook to a file
      const buffer = await workbook.xlsx.writeBuffer();

      return buffer;
    } catch (error) {
      console.error("export.downloadCSV: error writing buffer:", error);
    }
  },

  /**
   * Create the query to get the elements
   * @param {json} inputObject the base rules
   * @param {string} uid the table (i.e. api::product.product)
   * @param {number} limit how many elements get
   * @param {number} offset the first element to get
   * @returns {json} the target rules
   */
  async restructureObject(inputObject, uid, limit, offset) {
    let excel = strapi.config.get("excel");

    let where = {};

    // shall use localize data?
    if (excel?.config[uid]?.locale == "true") {
      where = {
        locale: "en",
      };
    }

    let orderBy = {
      id: "asc",
    };

    const restructuredObject = {
      select: inputObject.columns || "*",
      populate: {},
      where,
      orderBy,
      limit: limit,
      offset: offset,
    };

    // add any relationship - just one level deep
    for (const key in inputObject.relation) {
      restructuredObject.populate[key] = {
        select: inputObject.relation[key].column,
      };
    }

    return restructuredObject;
  },

  /**
   * Format an element
   * @param {json} data the element to format
   * @param {json} objectStructure the rules
   * @returns {json} the formatted element
   */
  async restructureData(data, objectStructure) {
    return data.map((item) => {
      const restructuredItem = {};

      // Restructure main data based on columns
      for (const key of objectStructure.columns) {
        if (key in item) {
          restructuredItem[key] = item[key];
        }
      }

      // Restructure relation data based on the specified structure
      for (const key in objectStructure.relation) {
        if (key in item) {
          const column = objectStructure.relation[key].column[0];
          if (item[key] && typeof item[key] === "object") {
            if (Array.isArray(item[key]) && item[key].length > 0) {
              restructuredItem[key] = item[key]
                .map((obj) => obj[column])
                .join(" ");
            } else {
              restructuredItem[key] = item[key][column];
            }
          } else {
            // Handle the case where item[key] is not an object
            restructuredItem[key] = null; // Or handle it as needed
          }
        }
      }

      return restructuredItem;
    });
  },
});
