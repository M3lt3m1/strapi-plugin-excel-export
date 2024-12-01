"use strict";
const ExcelJS = require("exceljs");
const excel = strapi.config.get("excel");

/**
 * Unflatten a flattened json where flattened names use "." as the delimiter and [INDEX] for arrays.
 * @param {json} data flattened json
 * @returns {json} un flattened json
 *
 * un-flattened | flattened
 * ---------------------------
 * {foo:{bar:false}} => {"foo.bar":false}
 * {a:[{b:["c","d"]}]} => {"a[0].b[0]":"c","a[0].b[1]":"d"}
 * [1,[2,[3,4],5],6] => {"[0]":1,"[1].[0]":2,"[1].[1].[0]":3,"[1].[1].[1]":4,"[1].[2]":5,"[2]":6}
 */
function unFlatten(data) {
    "use strict";
    if (Object(data) !== data || Array.isArray(data))
        return data;
    var result = {}, cur, prop, idx, last, temp;
    for(var p in data) {
        cur = result, prop = "", last = 0;
        do {
            idx = p.indexOf(".", last);
            temp = p.substring(last, idx !== -1 ? idx : undefined);
            cur = cur[prop] || (cur[prop] = (!isNaN(parseInt(temp)) ? [] : {}));
            prop = temp;
            last = idx + 1;
        } while(idx >= 0);
        cur[prop] = data[p];
    }
    return result[""];
}

/**
 * Flatten a json where flattened names use "." as the delimiter and [INDEX] for arrays.
 * @param {json} data the json
 * @returns {json} the flattened json
 *
 * un-flattened | flattened
 * ---------------------------
 * {foo:{bar:false}} => {"foo.bar":false}
 * {a:[{b:["c","d"]}]} => {"a[0].b[0]":"c","a[0].b[1]":"d"}
 * [1,[2,[3,4],5],6] => {"[0]":1,"[1].[0]":2,"[1].[1].[0]":3,"[1].[1].[1]":4,"[1].[2]":5,"[2]":6}
 */
function flatten(data,groupFlag=false) {
    var result = {};
    function recurse (cur, prop) {
        if (Object(cur) !== cur) {
            result[prop] = cur;
        } else if (Array.isArray(cur)) {
          if ( cur.length ) {
            if (groupFlag ) {
              result[prop] = Array.from(cur, el => {
                if ( typeof el === 'object') {
                  let text = '';
                  for ( let kk in el ) {
                    if( el[kk] !== 'object' ) {
                      text += '' + el[kk];
                    }
                  }
                  return text;
                }
                return el;
              }).join(', ')
            } else {
              for(var i=0, l=cur.length; i<l; i++)
                recurse(cur[i], prop ? prop+"."+i : ""+i, groupFlag);
              if (l == 0)
                result[prop] = [];
            }
          }
        } else {
            var isEmpty = true;
            for (var p in cur) {
                isEmpty = false;
                recurse(cur[p], prop ? prop+"."+p : p);
            }
            if (isEmpty)
                result[prop] = {};
        }
    }
    recurse(data, "");
    return result;
}

/**
 *
 * @param {json} jsonData
 * @param {json} _headers
 * @param {string} separator
 * @returns {string}
 */
function jsonToCsv(jsonData=[], _headers=null, separator='\t') {
  const array = typeof jsonData !== 'object' ? JSON.parse(jsonData) : jsonData;

  const csvRows = [];
  
  const headers = Array.from(Object.keys(_headers || array[0])).map( key => _headers[key] || key );
  csvRows.push(headers.join(separator));

  for (const row of array) {
      const values = headers.map(header => {
          let escaped = ('' + (row[header]||'')).replace(/"/g, '\\"');  // Escape delle virgolette
          escaped = escaped.replace(/[\t\r\n]/g, "");;             // rimuove cr lf tab
          return escaped.includes(separator) || escaped.includes('"') || escaped.includes('\n') ? `"${escaped}"` : escaped;
      });
      csvRows.push(values.join(separator));
  }

  return csvRows.join('\n');
}

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

    // strapi.log.info(`export.getTableData: query[${uid}]=${JSON.stringify(query, null, 2)}`);

    // query the data from the collection "uid"
    let response = await strapi.db.query(uid).findMany(query);

    // strapi.log.info(`export.getTableData: queryData[${uid}]=${JSON.stringify(response, null, 2).substring(0,3000)}...`);

    // build the header
    let headers = [
      ...excel?.config[uid]?.columns,                 // add field names of the collection
      ...Object.keys(excel?.config[uid]?.relations),  // add field names from relations of the collection
    ];

    // the name of the columns after flattening changes to the name of the labels
    headers = Object.keys(excel?.config[uid]?.labels||{});

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

    // format loaded data (i.e. remove unwanted fields, flatten remaining data)
    let tableData = await this.flattenData(response, excel?.config[uid]);

    // strapi.log.info(`export.getTableData: ${count} filtered records[${uid}]=${JSON.stringify(tableData, null, 2).substring(0,1000)}...`);

    let config = {
      count: count,
      labels: labelMap,
      columns: headers,
      data: tableData,
    };

    // strapi.log.info(`export.getTableData: output=${JSON.stringify(config, null, 2).substring(0,1000)}...`);

    return (config);
  },

  async downloadExcel(ctx) {
    try {

      // strapi.log.info(`export.downloadExcel: config=${JSON.stringify(excel, null, 2)}`);

      let uid = ctx?.query?.uid;

      // strapi.log.info(`export.downloadExcel: UID=${uid}`);

      let query = await this.makeQuery(excel?.config[uid], uid);

      // strapi.log.info(`export.downloadExcel: query=${JSON.stringify(query, null, 2)}`);

      let response = await strapi.db.query(uid).findMany(query);

      // strapi.log.info(`export.downloadExcel: response=${JSON.stringify(response, null, 2).substring(0,500)}...`);

      let excelData = await this.flattenData(response, excel?.config[uid]);

      // strapi.log.info(`export.downloadExcel: excelData=${JSON.stringify(excelData, null, 2).substring(0,500)}...`);

      // Create a new workbook and add a worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      // Extract column headers dynamically from the data
      let headers = [
        ...excel?.config[uid]?.columns,
        ...Object.keys(excel?.config[uid]?.relations),
      ];

      // the name of the columns after flattening changes to the name of the labels
      headers = Object.keys(excel?.config[uid]?.labels||{});

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

      // strapi.log.info(`export.downloadExcel: excelHeader=${JSON.stringify(headerRestructure, null, 2)}`);

      // Define dynamic column headers
      worksheet.columns = headers.map((header, index) => ({
        header: headerRestructure[index], // Use the formatted header
        key: header,
        width: 20,
      }));

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
      strapi.log.error("export.downloadExcel: ERROR "+error.toString());
    }
    return null;
  },

  async downloadCSV(ctx) {
    try {

      strapi.log.info(`export.downloadCSV: config=${JSON.stringify(excel, null, 2)}`);

      let uid = ctx?.query?.uid;

      strapi.log.info(`export.downloadCSV: UID=${uid}`);

      // make the query
      let query = await this.makeQuery(excel?.config[uid], uid);

      strapi.log.info(`export.downloadCSV: query=${JSON.stringify(query, null, 2)}`);

      // exec the query
      let response = await strapi.db.query(uid).findMany(query);

      strapi.log.info(`export.downloadCSV: response=${JSON.stringify(response, null, 2).substring(0,500)}...`);

      // filter and flatten the data
      let csvData = await this.flattenData(response, excel?.config[uid]);

      strapi.log.info(`export.downloadCSV: csvData=${JSON.stringify(csvData, null, 2).substring(0,500)}...`);

      // convert flattened data to CSV format
      let buffer = jsonToCsv(csvData,excel?.config[uid]?.labels?Object.keys(excel?.config[uid]?.labels):null);

      return buffer;
    } catch (error) {
      strapi.log.error("export.downloadCSV: ERROR "+error.toString());
    }
    return null;
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
      // strapi.log.info(`makeQuery.makeRelations: ${level}. ${owner} rules=${JSON.stringify(rules, null, 2)}`);
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
            strapi.log.error(`makeQuery.makeRelations: ${owner} ERROR cannot get key ${key}`);
          }
        }
      } catch(err) {
        strapi.log.error(`makeQuery.makeRelations: ${level}. ${owner} ERROR ${err.toString()}`);
      }
      // strapi.log.info(`makeQuery.makeRelations: ${level}. ${owner} relation=${JSON.stringify(relation, null, 2)}`);
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
   * @param {json} structureRules the rules
   * @returns {json} the formatted element
   */
  async flattenData(data, structureRules) {
    /**
     * Copy only the fields described in the rules
     * @param {json} item
     * @param {json} rules
     * @param {number} level
     * @returns {json} filtered item
     */
    const filterItem = (section, item, rules, level=1) => {
      let outItem = {};
      if ( !item ) return undefined;
      // strapi.log.info(`flattenData.filterItem: ${section} LEVEL ${level}${rules.mode?' '+rules.mode:''}, ${(rules.columns||[]).length} COLUMNS, ${Object.keys(rules.relations||{}).length} RELATIONS, INPUT=${JSON.stringify(item, null, 2)}\nRULES=${JSON.stringify(rules, null, 2)}\n`);
      for( let column of rules.columns||[]) {
        if ( item[column] ) {
          outItem[column] = item[column];
        }
      }
      for ( let key in rules.relations||{} ) {
        let subRules = rules.relations[key];
        let newItem  = item[key];
        if ( subRules && subRules ) {
          // strapi.log.info(`HANDLE ${key} ${typeof newItem}\n`);
          if( Array.isArray( newItem ) ) {
            outItem[key] = [];
            // strapi.log.info(`HANDLE ${key} array with ${newItem.length} columns items\n`);
            for ( let data of newItem ) {
                let ii = filterItem(key, data, subRules, level + 1);
                if ( ii ) {
                  outItem[key].push(ii);
                }
            }
            if ( rules.mode && rules.mode === 'group' ) {
              outItem[key] = outItem[key].join(rules.separator||', ')
            }
          } else if ( typeof newItem === 'object') {
            // strapi.log.info(`HANDLE ${key} FILTER ${(newItem.columns||[]).length} columns, ${Object.keys(subRules.relations||{}).length} relations\n`);
            outItem[key] = filterItem(key, newItem, subRules, level + 1);
          } else {
            // strapi.log.info(`HANDLE ${key} COPY ${newItem}`);
            outItem[key] = newItem;
          }
        } else {
          // strapi.log.info(`SKIP ${key} ${newItem?'':'NO DATA'} ${subRules?'':'NO RULES'}\n`);
        }
      }

      // strapi.log.info(`flattenData.filterItem: LEVEL ${level} OUTPUT=${JSON.stringify(outItem, null, 2)}\n`);
      if ( rules.mode && rules.mode === 'group' ) {
        let target = '';
        for ( let key in outItem ) {
          if ( typeof outItem[key] === 'object') {
            for ( let key2 in outItem[key] ) {
              target += outItem[key][key2]+(rules.separator||' '); // make it recursive
            }
          } else {
            target += outItem[key]+(rules.separator||' ');
          }
        }
        outItem = target.trim();
      }

      // strapi.log.info(`flattenData.filterItem: LEVEL ${level}, ${(rules.columns||[]).length} COLUMNS, ${Object.keys(rules.relations||{}).length} RELATIONS, OUTPUT=${JSON.stringify(outItem, null, 2)}\n`);
      return outItem;
    };

    return data.map((item) => {
      // strapi.log.info(`flattenData: INPUT ITEM ${JSON.stringify(item, null, 2)}\n`);
      item = filterItem('root', item, structureRules, 1);
      // strapi.log.info(`flattenData: FILTERED ITEM ${JSON.stringify(item, null, 2)}\n`);
      item = flatten(item, true);
      // strapi.log.info(`flattenData: FLATTENED ITEM ${JSON.stringify(item, null, 2)}\n`);
      return item;
    });

  }

});
