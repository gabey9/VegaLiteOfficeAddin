let currentHandler = null;

Office.onReady(() => {
  // Attach once Excel is ready
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Remove old handler if already attached
    if (currentHandler) {
      sheet.onChanged.remove(currentHandler);
      currentHandler = null;
    }

    // Add new handler
    currentHandler = sheet.onChanged.add(drawChart);
    await context.sync();

    // Initial draw
    await drawChart();
  });
});

async function drawChart() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();

      const values = usedRange.values;
      if (values.length < 2 || values[0].length < 2) {
        document.getElementById("chart").innerHTML =
          "<p>Please ensure at least 2 columns of data.</p>";
        return;
      }

      // Assume first row = headers
      const headers = values[0];
      const data = values.slice(1).map((row) => ({
        [headers[0]]: row[0],
        [headers[1]]: row[1],
      }));

      // Chart type from dropdown
      const chartType = document.getElementById("chartType").value;

      // Vega-Lite spec
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v5.json",
        description: "Dynamic chart from Excel data",
        data: { values: data },
        mark: chartType,
        encoding: {
          x: { field: headers[0], type: "quantitative" },
          y: { field: headers[1], type: "quantitative" },
        },
      };

      vegaEmbed("#chart", spec, { actions: false });
    });
  } catch (error) {
    console.error(error);
    document.getElementById("chart").innerHTML =
      "<p style='color:red;'>Error: " + error + "</p>";
  }
}