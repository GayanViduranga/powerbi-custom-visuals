"use strict";

import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import DataView = powerbi.DataView;
import DataViewTable = powerbi.DataViewTable;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import PrimitiveValue = powerbi.PrimitiveValue;

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./settings";

import "./../style/visual.less";

export class Visual implements IVisual {
    private target: HTMLElement;
    private viewport: HTMLDivElement;
    private scene: HTMLDivElement;
    private img: HTMLImageElement;
    private overlay: HTMLDivElement;
    private tooltip: HTMLDivElement;
    private message: HTMLDivElement;
    private toolbar: HTMLDivElement;
    private summaryBox: HTMLDivElement;

    private settings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private scale: number = 1;
    private panX: number = 0;
    private panY: number = 0;
    private isDragging: boolean = false;
    private dragStartX: number = 0;
    private dragStartY: number = 0;
    private startPanX: number = 0;
    private startPanY: number = 0;

    private tagSize: number = 14;
    private lastRows: PrimitiveValue[][] = [];
    private lastColumns: DataViewMetadataColumn[] = [];

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.target.innerHTML = "";
        this.target.style.width = "100%";
        this.target.style.height = "100%";
        this.target.style.position = "relative";
        this.target.style.overflow = "hidden";
        this.target.style.background = "#ffffff";
        this.target.style.fontFamily = "Segoe UI, sans-serif";

        this.viewport = document.createElement("div");
        this.viewport.style.width = "100%";
        this.viewport.style.height = "100%";
        this.viewport.style.position = "relative";
        this.viewport.style.overflow = "hidden";
        this.viewport.style.cursor = "grab";

        this.scene = document.createElement("div");
        this.scene.style.position = "absolute";
        this.scene.style.left = "0";
        this.scene.style.top = "0";
        this.scene.style.width = "100%";
        this.scene.style.height = "100%";
        this.scene.style.transformOrigin = "0 0";

        this.img = document.createElement("img");
        this.img.style.width = "100%";
        this.img.style.height = "100%";
        this.img.style.objectFit = "contain";
        this.img.style.display = "none";
        this.img.style.position = "absolute";
        this.img.style.left = "0";
        this.img.style.top = "0";
        this.img.draggable = false;

        this.overlay = document.createElement("div");
        this.overlay.style.position = "absolute";
        this.overlay.style.left = "0";
        this.overlay.style.top = "0";
        this.overlay.style.width = "100%";
        this.overlay.style.height = "100%";

        this.scene.appendChild(this.img);
        this.scene.appendChild(this.overlay);
        this.viewport.appendChild(this.scene);

        this.toolbar = document.createElement("div");
        this.toolbar.style.position = "absolute";
        this.toolbar.style.top = "10px";
        this.toolbar.style.right = "10px";

        this.summaryBox = document.createElement("div");
        this.summaryBox.style.position = "absolute";
        this.summaryBox.style.left = "10px";
        this.summaryBox.style.bottom = "10px";
        this.summaryBox.style.zIndex = "30";
        this.summaryBox.style.background = "rgba(255, 255, 255, 0.88)";
        this.summaryBox.style.border = "1px solid #d0d0d0";
        this.summaryBox.style.borderRadius = "6px";
        this.summaryBox.style.padding = "8px 10px";
        this.summaryBox.style.maxWidth = "320px";
        this.summaryBox.style.maxHeight = "180px";
        this.summaryBox.style.overflowY = "auto";
        this.summaryBox.style.fontSize = "12px";
        this.summaryBox.style.color = "#333";
        this.toolbar.style.display = "flex";
        this.toolbar.style.gap = "6px";
        this.toolbar.style.flexWrap = "wrap";
        this.toolbar.style.zIndex = "30";
        this.toolbar.style.maxWidth = "220px";

        this.tooltip = document.createElement("div");
        this.tooltip.style.position = "absolute";
        this.tooltip.style.display = "none";
        this.tooltip.style.minWidth = "160px";
        this.tooltip.style.maxWidth = "260px";
        this.tooltip.style.padding = "10px 12px";
        this.tooltip.style.borderRadius = "6px";
        this.tooltip.style.background = "rgba(34,34,34,0.96)";
        this.tooltip.style.color = "#ffffff";
        this.tooltip.style.fontSize = "12px";
        this.tooltip.style.lineHeight = "1.45";
        this.tooltip.style.pointerEvents = "none";
        this.tooltip.style.zIndex = "50";
        this.tooltip.style.boxShadow = "0 4px 14px rgba(0,0,0,0.25)";

        this.message = document.createElement("div");
        this.message.style.width = "100%";
        this.message.style.height = "100%";
        this.message.style.display = "flex";
        this.message.style.alignItems = "center";
        this.message.style.justifyContent = "center";
        this.message.style.color = "#666";
        this.message.style.fontSize = "14px";
        this.message.innerText = "Add Image URL, X, Y, Grade and Tooltip fields";

        this.target.appendChild(this.viewport);
        this.target.appendChild(this.toolbar);
        this.target.appendChild(this.summaryBox);
        this.target.appendChild(this.tooltip);
        this.target.appendChild(this.message);

        this.settings = new VisualFormattingSettingsModel();
        this.formattingSettingsService = new FormattingSettingsService();

        this.buildToolbar();
        this.initializeInteractions();
    }

    private createButton(label: string, onClick: () => void): HTMLButtonElement {
        const button = document.createElement("button");
        button.textContent = label;
        button.style.border = "1px solid #d0d0d0";
        button.style.background = "#ffffff";
        button.style.color = "#333";
        button.style.padding = "6px 10px";
        button.style.borderRadius = "4px";
        button.style.fontSize = "12px";
        button.style.cursor = "pointer";
        button.style.boxShadow = "0 1px 4px rgba(0,0,0,0.15)";
        button.style.pointerEvents = "auto";

        button.addEventListener("click", (e: MouseEvent) => {
            e.preventDefault();
            e.stopPropagation();
            onClick();
        });

        return button;
    }

    private buildToolbar(): void {
        this.toolbar.innerHTML = "";

        this.toolbar.appendChild(this.createButton("+", () => {
            this.zoomAtCenter(1.2);
        }));

        this.toolbar.appendChild(this.createButton("-", () => {
            this.zoomAtCenter(0.8);
        }));

        this.toolbar.appendChild(this.createButton("Reset", () => {
            this.scale = 1;
            this.panX = 0;
            this.panY = 0;
            this.applyTransform();
        }));

        this.toolbar.appendChild(this.createButton("A+", () => {
            this.tagSize = Math.min(6, this.tagSize + 0.25);
            this.renderMarkers();
        }));

        this.toolbar.appendChild(this.createButton("A-", () => {
            this.tagSize = Math.max(0.1, this.tagSize - 0.25);
            this.renderMarkers();
        }));
    }

    private zoomAtCenter(factor: number): void {
        const mouseX = this.viewport.clientWidth / 2;
        const mouseY = this.viewport.clientHeight / 2;

        const newScale = Math.max(0.5, Math.min(5, this.scale * factor));
        const worldX = (mouseX - this.panX) / this.scale;
        const worldY = (mouseY - this.panY) / this.scale;

        this.panX = mouseX - worldX * newScale;
        this.panY = mouseY - worldY * newScale;
        this.scale = newScale;

        this.applyTransform();
    }

    private initializeInteractions(): void {
        this.viewport.addEventListener("wheel", (event: WheelEvent) => {
            event.preventDefault();

            const rect = this.viewport.getBoundingClientRect();
            const mouseX = event.clientX - rect.left;
            const mouseY = event.clientY - rect.top;

            const zoomFactor = event.deltaY < 0 ? 1.1 : 0.9;
            const newScale = Math.max(0.5, Math.min(5, this.scale * zoomFactor));

            const worldX = (mouseX - this.panX) / this.scale;
            const worldY = (mouseY - this.panY) / this.scale;

            this.panX = mouseX - worldX * newScale;
            this.panY = mouseY - worldY * newScale;
            this.scale = newScale;

            this.applyTransform();
        });

        this.viewport.addEventListener("mousedown", (event: MouseEvent) => {
            const targetEl = event.target as HTMLElement;
            if (targetEl.tagName === "BUTTON") {
                return;
            }

            this.isDragging = true;
            this.dragStartX = event.clientX;
            this.dragStartY = event.clientY;
            this.startPanX = this.panX;
            this.startPanY = this.panY;
            this.viewport.style.cursor = "grabbing";
        });

        window.addEventListener("mousemove", (event: MouseEvent) => {
            if (!this.isDragging) {
                return;
            }

            this.panX = this.startPanX + (event.clientX - this.dragStartX);
            this.panY = this.startPanY + (event.clientY - this.dragStartY);
            this.applyTransform();
        });

        window.addEventListener("mouseup", () => {
            this.isDragging = false;
            this.viewport.style.cursor = "grab";
        });

        this.viewport.addEventListener("dblclick", () => {
            this.scale = 1;
            this.panX = 0;
            this.panY = 0;
            this.applyTransform();
        });
    }

    private applyTransform(): void {
        this.scene.style.transform = `translate(${this.panX}px, ${this.panY}px) scale(${this.scale})`;
    }

    private getGradeColor(grade: string): string {
        const value = (grade || "").toLowerCase().trim();

        if (value === "correct") {
            return "#28a745";
        }
        if (value === "tag_outside_of_mask") {
            return "#ff9800";
        }
        if (value === "missed_mask") {
            return "#2196f3";
        }
        if (value === "incorrect_product") {
            return "#f44336";
        }

        return "#9e9e9e";
    }

    private escapeHtml(value: string): string {
        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    private showTooltip(
        x: number,
        y: number,
        grade: string,
        tooltipItems: Array<{ name: string; value: string }>
    ): void {
        let html = `<div style="margin-bottom:6px;"><strong>Grade:</strong> ${this.escapeHtml(grade || "-")}</div>`;

        tooltipItems.forEach((item) => {
            html += `<div><strong>${this.escapeHtml(item.name)}:</strong> ${this.escapeHtml(item.value || "-")}</div>`;
        });

        this.tooltip.innerHTML = html;
        this.tooltip.style.display = "block";

        const left = Math.min(x + 14, this.target.clientWidth - 270);
        const top = Math.min(y + 14, this.target.clientHeight - 120);

        this.tooltip.style.left = `${Math.max(8, left)}px`;
        this.tooltip.style.top = `${Math.max(8, top)}px`;
    }

    private hideTooltip(): void {
        this.tooltip.style.display = "none";
    }

    private renderMarkers(): void {
        this.overlay.innerHTML = "";

        if (
            !this.lastRows ||
            this.lastRows.length === 0 ||
            !this.img.naturalWidth ||
            !this.img.naturalHeight
        ) {
            return;
        }

        const naturalWidth = this.img.naturalWidth;
        const naturalHeight = this.img.naturalHeight;
        const containerWidth = this.viewport.clientWidth;
        const containerHeight = this.viewport.clientHeight;

        const imageAspect = naturalWidth / naturalHeight;
        const containerAspect = containerWidth / containerHeight;

        let renderedWidth = 0;
        let renderedHeight = 0;
        let offsetX = 0;
        let offsetY = 0;

        if (imageAspect > containerAspect) {
            renderedWidth = containerWidth;
            renderedHeight = containerWidth / imageAspect;
            offsetY = (containerHeight - renderedHeight) / 2;
        } else {
            renderedHeight = containerHeight;
            renderedWidth = containerHeight * imageAspect;
            offsetX = (containerWidth - renderedWidth) / 2;
        }

        this.lastRows.forEach((row: PrimitiveValue[]) => {
            const x = Number(row[1]);
            const y = Number(row[2]);
            const grade = row[3]?.toString() || "";
            const color = this.getGradeColor(grade);

            if (isNaN(x) || isNaN(y)) {
                return;
            }

            const markerX = offsetX + (x / naturalWidth) * renderedWidth;
            const markerY = offsetY + (y / naturalHeight) * renderedHeight;

            const tooltipItems: Array<{ name: string; value: string }> = [];

            for (let i = 4; i < row.length; i++) {
                const columnName = this.lastColumns[i]?.displayName || `Field ${i - 3}`;
                const value = row[i] != null ? row[i]!.toString() : "";
                tooltipItems.push({ name: columnName, value });
            }

            const marker = document.createElement("div");
            marker.style.position = "absolute";
            marker.style.width = `${this.tagSize}px`;
            marker.style.height = `${this.tagSize}px`;
            marker.style.borderRadius = "50%";
            marker.style.background = color;
            marker.style.border = "none";
            marker.style.boxShadow = "0 0 0 1px rgba(0,0,0,0.25)";
            marker.style.left = `${markerX - this.tagSize / 2}px`;
            marker.style.top = `${markerY - this.tagSize / 2}px`;
            marker.style.cursor = "pointer";
            marker.style.pointerEvents = "auto";
            marker.style.zIndex = "5";

            marker.addEventListener("mouseenter", () => {
                const px = this.panX + markerX * this.scale;
                const py = this.panY + markerY * this.scale;
                this.showTooltip(px, py, grade, tooltipItems);
            });

            marker.addEventListener("mouseleave", () => {
                this.hideTooltip();
            });

            this.overlay.appendChild(marker);
        });
    }

    public update(options: VisualUpdateOptions): void {
        const dataView: DataView = options.dataViews?.[0];
        const table: DataViewTable | undefined = dataView?.table;

        this.settings = this.formattingSettingsService.populateFormattingSettingsModel(
            VisualFormattingSettingsModel,
            dataView
        );
        
        

        this.tagSize = Math.max(0.1, Number(this.settings.markerSettingsCard.tagSize.value || 0.8));

        this.hideTooltip();
        this.buildToolbar();

        if (!table || !table.rows || table.rows.length === 0) {
            this.lastRows = [];
            this.lastColumns = [];
            this.overlay.innerHTML = "";
            this.summaryBox.innerHTML = "";
            this.img.style.display = "none";
            this.message.style.display = "flex";
            return;
        }

        this.lastRows = table.rows;
        this.lastColumns = table.columns;

        const firstRow = table.rows[0];
        const imageUrl = firstRow[0]?.toString();

        if (!imageUrl) {
            this.overlay.innerHTML = "";
            this.img.style.display = "none";
            this.message.style.display = "flex";
            this.message.innerText = "Image URL is missing";
            return;
        }

        this.img.onerror = () => {
            this.overlay.innerHTML = "";
            this.img.style.display = "none";
            this.message.style.display = "flex";
            this.message.innerText = "Image failed to load";
        };

        this.img.onload = () => {
            this.renderMarkers();
            this.img.style.display = "block";
            this.message.style.display = "none";
        };

        this.renderSummary(this.lastRows);

        if (this.img.src !== imageUrl) {
            this.img.src = imageUrl;
        } else {
            this.renderMarkers();
            this.img.style.display = "block";
            this.message.style.display = "none";
        }
    }

    private renderSummary(rows: PrimitiveValue[][]): void {
        if (!rows || rows.length === 0) {
            this.summaryBox.innerHTML = "No data";
            return;
        }

        const gradeIndex = this.lastColumns.findIndex((col) => col.roles?.grade);
        const errorIndex = this.lastColumns.findIndex((col) => col.roles?.errors);

        if (errorIndex < 0) {
            this.summaryBox.innerHTML = "Error: map a numeric field into Errors role";
            return;
        }

        const chosenGradeIndex = gradeIndex >= 0 ? gradeIndex : 3;

        const errorSums: { [key: string]: number } = {};
        let totalErrorCount = 0;

        for (const row of rows) {
            const grade = (row[chosenGradeIndex] || "<empty>").toString();
            const rawValue = row[errorIndex];
            const errorValue = Number(rawValue);

            if (Number.isNaN(errorValue)) {
                this.summaryBox.innerHTML = `Error: value '${rawValue}' in Errors is not numeric`;
                return;
            }

            errorSums[grade] = (errorSums[grade] || 0) + errorValue;
            totalErrorCount += errorValue;
        }

        const sortedGrades = Object.keys(errorSums).sort((a, b) => errorSums[b] - errorSums[a]);

        let html = "<strong>Tag Summary</strong><br/>";
        html += `Total tag count: ${totalErrorCount}<br/><br/>`;

        sortedGrades.forEach((grade) => {
            html += `${this.escapeHtml(grade)}: ${errorSums[grade]}<br/>`;
        });

        this.summaryBox.innerHTML = html;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.settings);
    }
}
