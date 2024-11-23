"use strict";
const ExcelJS = require("exceljs");
const excel = strapi.config.get("excel");

module.exports = ({ strapi }) => ({

  // select the collections that are defined into the "excel" configuration
  async getDropDownData() {
    let dropDownValues = [];
    let configCollections = Object.keys(excel?.config);

    // scan all the collections
    strapi?.db?.config?.models?.forEach((element) => {
      if (element?.kind == "collectionType") {
        configCollections?.forEach((data) => {
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
    let uid = ctx?.query?.uid;
    let limit = ctx?.query?.limit;
    let offset = ctx?.query?.offset;
    let query = await this.makeQuery(excel?.config[uid], uid, limit, offset);

    strapi.log.info(`export.getTableData: query[${uid}]=${JSON.stringify(query, null, 2)}`);

    // query the data from the collection "uid"
    let response = await strapi.db.query(uid).findMany(query);

    strapi.log.info(`export.getTableData: queryData[${uid}]=${JSON.stringify(response, null, 2).substring(0,3000)}...`);

    // build the header
    let headers = [
      ...excel?.config[uid]?.columns,                 // add field names of the collection
      ...Object.keys(excel?.config[uid]?.relations),  // add field names from relations of the collection
    ];

    let labelMap = excel?.config[uid]?.labels||{}
    let labels = Array.from(headers, (name) => labelMap[name]||name)

    let where = {};

    if (excel?.config[uid]?.locale == "true") {
      where = {
        locale: "en",
      };
    }

    // find how many occurrences has been found with current query
    let count = await strapi.db.query(uid).count(where);

    // format loaded data
    let tableData = await this.flattenData(response, excel?.config[uid]);

    strapi.log.info(`export.getTableData: query[${uid}]=${JSON.stringify(response, null, 2).substring(0,1000)}...`);

    let config = {
      count: count,
      labels: labelMap,
      columns: headers,
      data: tableData,
    };

    strapi.log.info(`export.getTableData: output=${JSON.stringify(config, null, 2).substring(0,1000)}...`);

    return (config);
  },

  async downloadExcel(ctx) {
    try {

      strapi.log.info(`export.downloadExcel: config=${JSON.stringify(excel, null, 2)}`);

      let uid = ctx?.query?.uid;

      strapi.log.info(`export.downloadExcel: UID=${uid}`);

      let query = await this.makeQuery(excel?.config[uid], uid);

      strapi.log.info(`export.downloadExcel: query=${JSON.stringify(query, null, 2)}`);

      let response = await strapi.db.query(uid).findMany(query);

      strapi.log.info(`export.downloadExcel: response=${JSON.stringify(response, null, 2).substring(0,500)}...`);

      let excelData = await this.flattenData(response, excel?.config[uid]);

      strapi.log.info(`export.downloadExcel: excelData=${JSON.stringify(excelData, null, 2).substring(0,500)}...`);

      // Create a new workbook and add a worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      // Extract column headers dynamically from the data
      let headers = [
        ...excel?.config[uid]?.columns,
        ...Object.keys(excel?.config[uid]?.relations),
      ];
      let labelMap = excel?.config[uid]?.labels||{}
      let labels = Array.from(headers, (name) => labelMap[name]||name)

      // Transform the original headers to the desired format
      let headerRestructure = [];
      // headers?.forEach((element) => {
      labels?.forEach((element) => {
        const formattedHeader = element
          .split("_")
          .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
          .join(" ");
        headerRestructure.push(formattedHeader);
      });

      strapi.log.info(`export.downloadExcel: excelHeader=${JSON.stringify(headerRestructure, null, 2)}`);

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

      strapi.log.info(`export.downloadExcel: config=${JSON.stringify(excel, null, 2)}`);

      let uid = ctx?.query?.uid;

      strapi.log.info(`export.downloadCSV: UID=${uid}`);

      // make the query
      let query = await this.makeQuery(excel?.config[uid], uid);

      strapi.log.info(`export.downloadCSV: query=${JSON.stringify(query, null, 2)}`);

      // exec the query
      let response = await strapi.db.query(uid).findMany(query);

      strapi.log.info(`export.downloadCSV: response=${JSON.stringify(response, null, 2).substring(0,500)}...`);

      // format the data
      let excelData = await this.flattenData(response, excel?.config[uid]);

      strapi.log.info(`export.downloadCSV: excelData=${JSON.stringify(excelData, null, 2).substring(0,500)}...`);

      // Create a new workbook and add a worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      // Extract column headers dynamically from the data
      let headers = [
        ...excel?.config[uid]?.columns,
        ...Object.keys(excel?.config[uid]?.relations),
      ];
      let labelMap = excel?.config[uid]?.labels||{}
      let labels = Array.from(headers, (name) => labelMap[name]||name)

      // Transform the original headers to the desired format
      let headerRestructure = [];
      // headers?.forEach((element) => {
      labels?.forEach((element) => {
        const formattedHeader = element
          .split("_")
          .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
          .join(" ");
        headerRestructure.push(formattedHeader);
      });

      strapi.log.info(`export.downloadCSV: excelHeader=${JSON.stringify(headerRestructure, null, 2)}`);

      // Define dynamic column headers
      worksheet.columns = headers.map((header, index) => ({
        header: headerRestructure[index], // Use the formatted header
        key: header,
        width: 20,
      }));

      // Define the dropdown list options for the Gender column

      // Add data to the worksheet, row by row
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
   * @param {json} collectionCfg the rules for the collection
   * @param {string} uid collection unique id (i.e. api::product.product)
   * @param {number} limit how many elements get
   * @param {number} offset the first element to get
   * @returns {json} the target rules
   */
  async makeQuery(collectionCfg, uid, limit, offset) {
    /**
     * Build recursively the relations
     * @param {json} rules the rules to use
     * EXAMPLE:
     *
     * input rules = {
     *   columns: ["pippo","pluto"],
     *   relations: {
     *     prices: {
     *       columns: ["unit_cost"],
     *       relations: {
     *         supplier: {
     *           columns: ["name"]
     *         }
     *       }
     *     }
     *   }
     * },
     *
     * expected output = {
     *  "select": [
     *    "pippo",
     *    "pluto",
     *  ],
     *  "populate": {
     *    "prices": {
     *      "select": [
     *        "unit_cost"
     *      ],
     *      "populate": {
     *         "supplier": {
     *            "select": [
     *               "name"
     *            ],
     *         }
     *      }
     *    },
     */
    let makeRelations = (rules, owner, level) => {
      let relation = {};
      // console.log(`makeQuery.makeRelations: ${level}. ${owner} rules=${JSON.stringify(rules, null, 2)}`);
      try {
        // any columns to select?
        if ( rules.columns && rules.columns.length ) {
          relation.select = rules.columns;
        }
        // any relation to use?
        for (const key in rules.relations||[]) {
          if ( rules.relations[key] ) {
            relation.populate = relation.populate || {};
            relation.populate[key] = makeRelations(rules.relations[key], key, level + 1); // add a new relation
          } else {
            console.error(`makeQuery.makeRelations: ${owner} ERROR cannot get key ${key}`);
          }
        }
      } catch(err) {
        console.error(`makeQuery.makeRelations: ${level}. ${owner} ERROR ${err.toString()}`);
      }
      // console.log(`makeQuery.makeRelations: ${level}. ${owner} relation=${JSON.stringify(relation, null, 2)}`);
      return relation;
    };
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

    const query = {
      select: collectionCfg.columns || "*",
      populate: {},
      where,
      orderBy,
      limit: limit,
      offset: offset,
    };

    // add any relationship - any level deep
    for (const key in collectionCfg.relations) {
      // columns or relations defined?
      if ( collectionCfg.relations[key]?.columns.length || Object.keys(collectionCfg.relations[key]?.relations||{}).length > 0){
        query.populate[key] = makeRelations(collectionCfg.relations[key], key, 1);
      }
    }

    return query;
  },

  /**
   * Flatten each element of the provided collection
   * @param {json} data the elements to flatten
   * @param {json} objectStructure the rules
   * @returns {json} the formatted element
   */
  async flattenData(data, objectStructure) {
    return data.map((item) => {
      const flattenItem = {};

      // flatten main data based on columns
      for (const key of objectStructure.columns) {
        if (key in item) {
          flattenItem[key] = item[key];
        }
      }

      // flatten relations data based on the specified structure - shall be recursive
      for (const key in objectStructure.relations) {
        if (key in item) {
          const column = objectStructure.relations[key].columns[0];
          if (item[key] && typeof item[key] === "object") {
            if (Array.isArray(item[key]) && item[key].length > 0) {
              flattenItem[key] = item[key]
                .map((obj) => obj[column])
                .join(", ");
            } else {
              flattenItem[key] = item[key][column];
            }
          } else {
            // Handle the case where item[key] is not an object
            flattenItem[key] = null; // Or handle it as needed
          }
        }
      }

      return flattenItem;
    });
  },
});
