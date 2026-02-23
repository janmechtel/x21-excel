/* import { Tool } from "../../types/index.ts";
import { ExcelApiConfigService } from "../../utils/excel-api-config.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger('CreateChartTool');

export class CreateChartTool implements Tool {
  name = "create_chart";
  description = "Create a chart in Excel from specified data range.";
  input_schema = {
    worksheet: { type: "string", description: "The worksheet name" },
    dataRange: { type: "string", description: "The data range for the chart" },
    chartType: { type: "string", description: "Type of chart to create" },
    title: { type: "string", description: "Chart title (optional)" }
  };

  async execute(params: any, workbookName: string): Promise<any> {
    logger.info("🔧 create_chart called with params:", params, "workbookName:", workbookName);

    // Validate required fields
    if (!params.worksheet || !params.dataRange || !params.chartType || !workbookName) {
      throw new Error("worksheet, dataRange, chartType, and workbookName are required");
    }

    logger.info("📡 Making HTTP call to Excel API for create_chart...");

    try {
      const excelApi = ExcelApiConfigService.getInstance();
      const response = await fetch(excelApi.getApiUrlActionExecution(), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          action: "create_chart",
          worksheet: params.worksheet,
          dataRange: params.dataRange,
          chartType: params.chartType,
          workbookName: workbookName,
          title: params.title || "",
          position: params.position || { left: 100, top: 100, width: 400, height: 300 }
        }),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json();
      logger.info("✅ create_chart completed, chart created via Excel API");

      return result;
    } catch (error) {
      logger.error("❌ Error calling Excel API:", error);
      throw error;
    }
  }
} */
