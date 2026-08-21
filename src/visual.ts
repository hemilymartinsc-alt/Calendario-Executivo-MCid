import powerbi from "powerbi-visuals-api";
import RuntimeVisual from "./runtime";
import "./../style/visual.less";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

/**
 * Thin TypeScript entry point. All visual behavior lives in runtime.js and is
 * bundled by the official Power BI Visuals Tools pipeline.
 */
export class Visual extends (RuntimeVisual as any) implements IVisual {
    constructor(options: VisualConstructorOptions) {
        super(options);
    }

    public update(options: VisualUpdateOptions): void {
        super.update(options);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return super.getFormattingModel();
    }
}
