import React, { useEffect, useState } from "react";
import DataTable from "react-data-table-component";

import axios from "axios";
import {
  Box,
  ContentLayout,
  Button,
  HeaderLayout,
  Layout,
  Combobox,
  ComboboxOption,
  Stack,
  Typography,
} from "@strapi/design-system";

const HomePage = () => {
  const baseUrl = process.env.STRAPI_ADMIN_BACKEND_URL;

  // which collection to export
  const [dropDownData, setDropDownData] = useState([]);
  const [selectedValue, setSelectedValue] = useState(null);

  // which export format to use
  const [exportFormats, setExportFormats] = useState([{label:'cvs', value:'csv'},{label:'xlsx',value:'xlsx'}]);   // which format to use
  const [selectedFormat, setSelectedFormat] = useState('cvs');

  const [labels, setLabels] = useState([]);
  const [columns, setColumns] = useState([]);
  const [tableData, setTableData] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSuccessMessage, setIsSuccessMessage] = useState(false);
  const [fileName, setFileName] = useState("");

  //data table pagination
  const [loading, setLoading] = useState(false);
  const [totalRows, setTotalRows] = useState(0);
  const [perPage, setPerPage] = useState(10);

  useEffect(() => {
    const fetchData = async () => {
      try {
        const response = await axios.get(`${baseUrl}/excel-export/get/dropdown/values`);
        setDropDownData(response.data);
        setIsLoading(false);
      } catch (error) {
        console.error("HomePage.useEffect: Error fetching dropdown values:", error);
        setIsLoading(false);
      }
    };

    fetchData();
  }, []);

  //data table pagination
  const handleComboboxChange = async (value) => {
    setSelectedValue(value); // Use the callback form to ensure state is updated
    if (value) {
      fetchPageData(value, 1, 10);
    }
  };

  //data table pagination
  const handleOutputFormatChange = async (value) => {
    console.log(`handleOutputFormatChange: set ${value}`);
    setSelectedFormat(value); // Use the callback form to ensure state is updated
  };

  /**
   * Run the download process
   */
  const handleDownload = async () => {
    if ( selectedFormat === 'xslx' ) {
      await handleDownloadExcel();
    }
    await handleDownloadCSV();
  }

  /**
   * Download data in Excel format
   */
  const handleDownloadExcel = async () => {
    try {
      let ff = selectedValue.split('.');
      let name = ff[ff.length - 1];
      console.log(`handleDownloadExcel: RUN ${selectedValue} => ${name}`);
      const response = await axios.get(`${baseUrl}/excel-export/download/excel`,
        {
          responseType: "arraybuffer",
          params: {
            uid: selectedValue,
          },
        }
      );

      // Create a Blob from the response data and trigger download
      if (response.data) {
        const currentDate = new Date();
        const formattedDate = formatDate(currentDate);
        setFileName(`${name}-${formattedDate}.xlsx`);

        const blob = new Blob([response.data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const link = document.createElement("a");
        link.href = window.URL.createObjectURL(blob);
        link.download = `${name}-${formattedDate}.xlsx`;
        link.click();
        setIsSuccessMessage(true);
        // Hide the success message after 3000 milliseconds (3 seconds)
        setTimeout(() => {
          setIsSuccessMessage(false);
        }, 8000);
      }
    } catch (error) {
      console.error("handleDownloadExcel: Error downloading Excel file:", error);
    }
  };


  /**
   * Download data in CSV format
   */
  const handleDownloadCSV = async () => {
    try {
      let ff = selectedValue.split('.');
      let name = ff[ff.length - 1];
      console.log(`handleDownloadCSV: RUN ${selectedValue} => ${name}`);
      const response = await axios.get(
        `${baseUrl}/excel-export/download/csv`,
        {
          responseType: "arraybuffer",
          params: {
            uid: selectedValue,
          },
        }
      );

      // Create a Blob from the response data and trigger download
      if (response.data) {
        const currentDate = new Date();
        const formattedDate = formatDate(currentDate);
        setFileName(`${name}-${formattedDate}.csv`);
        const blob = new Blob([response.data], {
          type: "application/text",
        });
        const link = document.createElement("a");
        link.href = window.URL.createObjectURL(blob);
        link.download = `${name}-${formattedDate}.csv`;
        link.click();
        setIsSuccessMessage(true);
        // Hide the success message after 3000 milliseconds (3 seconds)
        setTimeout(() => {
          setIsSuccessMessage(false);
        }, 8000);
      }
    } catch (error) {
      console.error("handleDownloadCSV: Error downloading Excel file:", error);
    }
  };

  const handleComboBoxClear = async () => {
    setSelectedValue(null);
    setTableData([]);
  };

  const handleOutputFormatClear = async () => {
    setSelectedFormat('csv');
  };

  // name is label key or column name, selector is data hash
  const columnRestructure = columns.map((property) => ({
    name: labels ? labels[property] : property?.charAt(0).toUpperCase() + property?.slice(1).replace(/_/g, " "),
    selector: (row) => row[property],
  }));

  // Function to format date as "DD-MM-YYYY-HH-mm-ss"
  const formatDate = (date) => {
    const day = date.getDate().toString().padStart(2, "0");
    const month = (date.getMonth() + 1).toString().padStart(2, "0");
    const year = date.getFullYear();
    const hours = date.getHours().toString().padStart(2, "0");
    const minutes = date.getMinutes().toString().padStart(2, "0");
    const seconds = date.getSeconds().toString().padStart(2, "0");
    return `${day}-${month}-${year}-${hours}-${minutes}-${seconds}`;
  };

  // data table functionality

  /**
   *
   * @param {string} value collection
   * @param {number} page
   * @param {number} itemsPerPage
   */
  const fetchPageData = async (value, page, itemsPerPage) => {
    setLoading(true);
    const currentSelectedValue = value; // Store the selectedValue in a variable

    console.log(`fetchPageData: columnRestructure=${JSON.stringify(columnRestructure, null, 2)}`);

    if (currentSelectedValue) {
      try {
        const offset = (page - 1) * itemsPerPage; // Calculate the offset based on the current page and items per page
        const limit = itemsPerPage;

        const response = await axios.get(`${baseUrl}/excel-export/get/table/data?uid=${value}&limit=${limit}&offset=${offset}`);

        if (response?.data?.columns) {
          console.log("fetchPageData: set columns");
          setColumns(response.data.columns);
        }

        if (response?.data?.labels) {
          console.log("fetchPageData: set labels");
          setLabels(response.data.labels);
        }

        if (response?.data?.data) {
          setTableData(response.data.data);
          setTotalRows(response.data.count);
        }

      } catch (error) {
        console.error(`fetchPageData: error fetching page ${page} data: ${error.toString()}`);
      } finally {
        setLoading(false);
      }
    }
  };

  const handlePageChange = (page) => {
    fetchPageData(selectedValue, page, perPage);
  };

  /**
   * Load the data for the required page
   * @param {number} itemsPerPage
   * @param {number} currentPage
   */
  const handlePerRowsChange = async (itemsPerPage, currentPage) => {
    setLoading(true);
    try {
      const offset = (currentPage - 1) * itemsPerPage; // Calculate the offset based on the current page and items per page
      const limit = itemsPerPage;

      console.log(`handlePerRowsChange: columnRestructure=${JSON.stringify(columnRestructure, null, 2)}`);

      const response = await axios.get(`${baseUrl}/excel-export/get/table/data?uid=${selectedValue}&limit=${limit}&offset=${offset}`);

      if (response?.data?.labels) {
        console.log("handlePerRowsChange: set labels");
        setLabels(response.data.labels);
      }

      if (response?.data?.data) {
        setTableData(response.data.data);
        setPerPage(itemsPerPage);
      }

    } catch (error) {
      console.error("handlePerRowsChange: Error fetching table data:", error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Box background="neutral100">
      <Layout>
        <>
          <HeaderLayout title="Matchbox Export" as="h2" />
          <ContentLayout>
            <Stack>

              <Box padding={4} width="600px">
                <Combobox
                  label="Collection Type"
                  size="M"
                  onChange={handleComboboxChange}
                  value={selectedValue}
                  placeholder="Select collection type"
                  onClear={handleComboBoxClear}
                >
                  {dropDownData?.data?.map((item) => (
                    <ComboboxOption key={item.value} value={item.value}>
                      {item.label}
                    </ComboboxOption>
                  ))}
                </Combobox>
              </Box>

              <Box padding={4} width="600px">
                <Combobox
                  label="Output Format Type"
                  size="M"
                  onChange={handleOutputFormatChange}
                  value={selectedValue}
                  placeholder="Select output format type"
                  onClear={handleOutputFormatClear}
                >
                  {exportFormats?.map((item) => (
                    <ComboboxOption key={item.value} value={item.value}>
                      {item.label}
                    </ComboboxOption>
                  ))}
                </Combobox>
              </Box>

              {selectedValue && selectedFormat && (
                <>
                  <Box padding={4} marginTop={2} className="ml-auto">
                    <Button
                      size="L"
                      variant="default"
                      onClick={handleDownload}
                    >
                      Download
                    </Button>

                    <br />

                    {isSuccessMessage && (
                      <Typography
                        style={{
                          color: "green",
                          "font-size": "medium",
                          "font-weight": "500",
                        }}
                      >
                        Download completed: {fileName} successfully downloaded!
                      </Typography>
                    )}

                  </Box>

                  <Box className="ml-auto">
                    <DataTable
                      pagination
                      columns={columnRestructure}
                      data={tableData}
                      progressPending={loading}
                      paginationServer
                      paginationTotalRows={totalRows}
                      onChangeRowsPerPage={handlePerRowsChange}
                      onChangePage={handlePageChange}
                    />
                  </Box>

                </>
              )}
            </Stack>
          </ContentLayout>
        </>
      </Layout>
    </Box>
  );
};

export default HomePage;
