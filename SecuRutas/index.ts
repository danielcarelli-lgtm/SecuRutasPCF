import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as atlas from "azure-maps-control";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { ComboBox, IComboBoxOption, IComboBox, DatePicker, Dropdown, IDropdownOption, IconButton, initializeIcons } from "@fluentui/react";

initializeIcons(undefined, { disableWarnings: true });

interface RoutePoint {
    lng: number;
    lat: number;
    timestamp?: string;
}

export class SecuRutas implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _mapContainer: HTMLDivElement;
    private _headerContainer: HTMLDivElement;

    private _map: atlas.Map | undefined;
    private _satelliteTileLayer: atlas.layer.TileLayer | undefined;
    private _baseTileLayer: atlas.layer.TileLayer | undefined;

    private _routeSvg: SVGSVGElement | undefined;
    private _routeHalo: SVGPolylineElement | undefined;
    private _routeLine: SVGPolylineElement | undefined;
    private _waypointsGroup: SVGGElement | undefined;
    private _waypointCircles: SVGCircleElement[] = [];
    private _startMarker: SVGGElement | undefined;
    private _endMarker: SVGGElement | undefined;
    private _tooltipEl: HTMLDivElement | undefined;

    private _routePoints: RoutePoint[] = [];
    private _routePositions: atlas.data.Position[] = [];
    private _isMapReady = false;
    private _resizeObserver: ResizeObserver | undefined;
    private _pendingPoints: RoutePoint[] | undefined;

    private _selectedDate: Date = new Date();
    private _selectedResource: string | undefined;
    private _selectedMapStyle = "road";
    private _resources: IComboBoxOption[] = [];

    private _status = "";
    private _mapError = "";
    private _loading = false;

    private _webAPI: ComponentFramework.WebApi;
    private _root: Root;
    private readonly _controlVersion = "v1.1.33";

    private readonly _pageSize = 5000;
    private readonly _maxPages = 20;

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._container = container;
        this._webAPI = context.webAPI;

        const mainDiv = document.createElement("div");
        mainDiv.className = "secu-rutas-container";
        mainDiv.style.display = "flex";
        mainDiv.style.flexDirection = "column";
        mainDiv.style.width = "100%";
        mainDiv.style.height = "100%";
        mainDiv.style.minHeight = "500px";

        this._headerContainer = document.createElement("div");
        this._headerContainer.className = "secu-rutas-header";

        this._mapContainer = document.createElement("div");
        this._mapContainer.className = "secu-rutas-map";
        this._mapContainer.style.flex = "1 1 auto";
        this._mapContainer.style.minHeight = "400px";
        this._mapContainer.style.position = "relative";

        mainDiv.appendChild(this._headerContainer);
        mainDiv.appendChild(this._mapContainer);

        const versionDiv = document.createElement("div");
        versionDiv.className = "secu-rutas-version";
        versionDiv.innerText = this._controlVersion;
        mainDiv.appendChild(versionDiv);

        this._container.appendChild(mainDiv);

        this.renderHeader();
        this.initMap();
        void this.loadResources();
    }

    private initMap(): void {
        const azureMapsKey = this._context.parameters.azureMapsKey.raw;
        if (!azureMapsKey) {
            this.setStatus("⚠ azureMapsKey vacío: el mapa no se puede inicializar.");
            return;
        }

        atlas.setDomain("eu.atlas.microsoft.com");

        const map = new atlas.Map(this._mapContainer, {
            authOptions: {
                authType: atlas.AuthenticationType.subscriptionKey,
                subscriptionKey: azureMapsKey
            },
            domain: "eu.atlas.microsoft.com",
            view: "Auto",
            language: "es-ES",
            center: [-3.7038, 40.4168],
            zoom: 10,
            style: "blank"
        });
        this._map = map;

        map.events.add("error", (e: unknown) => {
            console.error("[SecuRutas] Error del mapa:", e);
            const inner = (e as { error?: { message?: string } })?.error;
            const msg = inner?.message ?? this.errMsg(e);
            // Ignorar errores internos de bing-mvt si ocurrieran
            if (msg.indexOf("bing-mvt") >= 0) return;
            this._mapError = "⚠ Azure Maps: " + msg;
            this.renderHeader();
        });

        map.events.add("ready", () => {
            this.setupRasterLayers(map, azureMapsKey);
            this.createRouteOverlay();

            map.events.add("render", () => this.drawRoute());
            map.events.add("movestart", () => this.hideTooltip());

            this._isMapReady = true;
            this.safeResize();

            if (this._pendingPoints) {
                const pending = this._pendingPoints;
                this._pendingPoints = undefined;
                this.updateMap(pending);
            }
        });

        if (typeof ResizeObserver !== "undefined") {
            this._resizeObserver = new ResizeObserver(() => this.safeResize());
            this._resizeObserver.observe(this._mapContainer);
        }
    }

    private getAzureTileUrl(tilesetId: string, azureMapsKey: string): string {
        return "https://eu.atlas.microsoft.com/map/tile?api-version=2.1" +
            "&tilesetId=" + tilesetId +
            "&zoom={z}&x={x}&y={y}&tileSize=256&language=es-ES&view=Auto" +
            "&subscription-key=" + encodeURIComponent(azureMapsKey);
    }

    private setupRasterLayers(map: atlas.Map, azureMapsKey: string): void {
        // Capa base de imágenes de satélite (solo visible en modo satélite)
        this._satelliteTileLayer = new atlas.layer.TileLayer({
            tileUrl: this.getAzureTileUrl("microsoft.imagery", azureMapsKey),
            tileSize: 256,
            maxSourceZoom: 18,
            visible: false
        });

        // Capa principal con nombres de calles, sentidos, parques y POIs renderizados
        this._baseTileLayer = new atlas.layer.TileLayer({
            tileUrl: this.getAzureTileUrl("microsoft.base.road", azureMapsKey),
            tileSize: 256,
            maxSourceZoom: 18,
            visible: true
        });

        map.layers.add([this._satelliteTileLayer, this._baseTileLayer]);
        this.applyMapStyleLayers();
    }

    private applyMapStyleLayers(): void {
        const azureMapsKey = this._context.parameters.azureMapsKey.raw;
        if (!azureMapsKey || !this._baseTileLayer || !this._satelliteTileLayer) return;

        if (this._selectedMapStyle === "satellite_road_labels") {
            this._satelliteTileLayer.setOptions({ visible: true });
            this._baseTileLayer.setOptions({
                tileUrl: this.getAzureTileUrl("microsoft.base.hybrid.road", azureMapsKey),
                tileSize: 256,
                maxSourceZoom: 18,
                visible: true
            });
        } else if (this._selectedMapStyle === "osm") {
            this._satelliteTileLayer.setOptions({ visible: false });
            this._baseTileLayer.setOptions({
                tileUrl: "https://tile.openstreetmap.org/{z}/{x}/{y}.png",
                tileSize: 256,
                maxSourceZoom: 19,
                visible: true
            });
        } else {
            this._satelliteTileLayer.setOptions({ visible: false });
            this._baseTileLayer.setOptions({
                tileUrl: this.getAzureTileUrl("microsoft.base.road", azureMapsKey),
                tileSize: 256,
                maxSourceZoom: 18,
                visible: true
            });
        }
    }

    private safeResize(): void {
        if (!this._isMapReady || !this._map) return;
        try {
            this._map.resize();
        } catch (e) {
            console.warn("[SecuRutas] resize ignorado:", e);
        }
    }

    private updateMap(points: RoutePoint[]): void {
        if (!this._isMapReady || !this._map) {
            this._pendingPoints = points;
            return;
        }

        this.hideTooltip();
        this._routePoints = points;
        this._routePositions = points.map(p => [p.lng, p.lat]);

        this.rebuildWaypointNodes();
        this.drawRoute();

        if (points.length === 0) return;

        let minLat = Infinity, maxLat = -Infinity, minLng = Infinity, maxLng = -Infinity;
        for (const p of points) {
            if (p.lng < minLng) minLng = p.lng;
            if (p.lng > maxLng) maxLng = p.lng;
            if (p.lat < minLat) minLat = p.lat;
            if (p.lat > maxLat) maxLat = p.lat;
        }

        this.safeResize();

        if (Math.abs(maxLat - minLat) > 0.0001 || Math.abs(maxLng - minLng) > 0.0001) {
            this._map.setCamera({
                bounds: atlas.data.BoundingBox.fromEdges(minLng, minLat, maxLng, maxLat),
                padding: 50
            });
        } else {
            this._map.setCamera({ center: [points[0].lng, points[0].lat], zoom: 15 });
        }
    }

    private createRouteOverlay(): void {
        const svgNamespace = "http://www.w3.org/2000/svg";
        const svg = document.createElementNS(svgNamespace, "svg");
        svg.setAttribute("aria-hidden", "true");
        svg.style.position = "absolute";
        svg.style.inset = "0";
        svg.style.width = "100%";
        svg.style.height = "100%";
        svg.style.overflow = "hidden";
        svg.style.pointerEvents = "none";
        svg.style.zIndex = "2";

        const halo = document.createElementNS(svgNamespace, "polyline");
        halo.setAttribute("fill", "none");
        halo.setAttribute("stroke", "#ffffff");
        halo.setAttribute("stroke-width", "8");
        halo.setAttribute("stroke-linecap", "round");
        halo.setAttribute("stroke-linejoin", "round");
        halo.setAttribute("vector-effect", "non-scaling-stroke");

        const line = document.createElementNS(svgNamespace, "polyline");
        line.setAttribute("fill", "none");
        line.setAttribute("stroke", "#004b8b");
        line.setAttribute("stroke-width", "5");
        line.setAttribute("stroke-linecap", "round");
        line.setAttribute("stroke-linejoin", "round");
        line.setAttribute("vector-effect", "non-scaling-stroke");

        const waypointsGroup = document.createElementNS(svgNamespace, "g") as SVGGElement;

        const startMarker = this.createEndpointMarker(svgNamespace, "#16834b", "start");
        const endMarker = this.createEndpointMarker(svgNamespace, "#c43131", "end");

        // Eventos de tooltip para Inicio y Fin
        startMarker.addEventListener("mouseenter", () => {
            if (this._routePoints.length > 0) {
                this.showTooltip(this._routePoints[0], 0, this._routePoints.length);
            }
        });
        startMarker.addEventListener("mouseleave", () => this.hideTooltip());

        endMarker.addEventListener("mouseenter", () => {
            if (this._routePoints.length > 1) {
                const lastIdx = this._routePoints.length - 1;
                this.showTooltip(this._routePoints[lastIdx], lastIdx, this._routePoints.length);
            }
        });
        endMarker.addEventListener("mouseleave", () => this.hideTooltip());

        svg.appendChild(halo);
        svg.appendChild(line);
        svg.appendChild(waypointsGroup);
        svg.appendChild(startMarker);
        svg.appendChild(endMarker);

        // Tooltip flotante HTML
        const tooltip = document.createElement("div");
        tooltip.style.position = "absolute";
        tooltip.style.display = "none";
        tooltip.style.pointerEvents = "none";
        tooltip.style.zIndex = "10";
        tooltip.style.backgroundColor = "rgba(255, 255, 255, 0.97)";
        tooltip.style.color = "#201f1e";
        tooltip.style.padding = "8px 10px";
        tooltip.style.borderRadius = "6px";
        tooltip.style.boxShadow = "0 2px 8px rgba(0, 0, 0, 0.28)";
        tooltip.style.border = "1px solid #edebe9";
        tooltip.style.fontFamily = "'Segoe UI', sans-serif";
        tooltip.style.fontSize = "12px";
        tooltip.style.lineHeight = "1.4";
        tooltip.style.whiteSpace = "nowrap";

        this._mapContainer.appendChild(svg);
        this._mapContainer.appendChild(tooltip);

        this._routeSvg = svg;
        this._routeHalo = halo;
        this._routeLine = line;
        this._waypointsGroup = waypointsGroup;
        this._startMarker = startMarker;
        this._endMarker = endMarker;
        this._tooltipEl = tooltip;
    }

    private createEndpointMarker(svgNamespace: string, color: string, kind: "start" | "end"): SVGGElement {
        const group = document.createElementNS(svgNamespace, "g") as SVGGElement;
        group.setAttribute("aria-hidden", "true");
        group.style.filter = "drop-shadow(0 1px 2px rgba(0,0,0,0.45))";
        group.style.pointerEvents = "all";
        group.style.cursor = "pointer";

        const circle = document.createElementNS(svgNamespace, "circle");
        circle.setAttribute("r", "13");
        circle.setAttribute("fill", color);
        circle.setAttribute("stroke", "#ffffff");
        circle.setAttribute("stroke-width", "2");
        group.appendChild(circle);

        const icon = document.createElementNS(svgNamespace, "path");
        icon.setAttribute("fill", "none");
        icon.setAttribute("stroke", "#ffffff");
        icon.setAttribute("stroke-width", "2");
        icon.setAttribute("stroke-linecap", "round");
        icon.setAttribute("stroke-linejoin", "round");
        icon.setAttribute(
            "d",
            kind === "start"
                ? "M -2 -6 L 6 0 L -2 6 Z"
                : "M -3 7 L -3 -7 M -2 -6 L 5 -4 L -2 -1"
        );
        if (kind === "start") icon.setAttribute("fill", "#ffffff");
        group.appendChild(icon);

        return group;
    }

    private rebuildWaypointNodes(): void {
        if (!this._waypointsGroup) return;

        while (this._waypointsGroup.firstChild) {
            this._waypointsGroup.removeChild(this._waypointsGroup.firstChild);
        }
        this._waypointCircles = [];

        if (this._routePoints.length <= 2) return;

        const svgNamespace = "http://www.w3.org/2000/svg";
        const total = this._routePoints.length;

        for (let i = 1; i < total - 1; i++) {
            const point = this._routePoints[i];
            const circle = document.createElementNS(svgNamespace, "circle");
            circle.setAttribute("r", "4.5");
            circle.setAttribute("fill", "#ffffff");
            circle.setAttribute("stroke", "#004b8b");
            circle.setAttribute("stroke-width", "2");
            circle.style.pointerEvents = "all";
            circle.style.cursor = "pointer";
            circle.style.filter = "drop-shadow(0 1px 1px rgba(0,0,0,0.3))";

            circle.addEventListener("mouseenter", () => {
                circle.setAttribute("r", "7");
                circle.setAttribute("fill", "#e8f2fc");
                this.showTooltip(point, i, total);
            });

            circle.addEventListener("mouseleave", () => {
                circle.setAttribute("r", "4.5");
                circle.setAttribute("fill", "#ffffff");
                this.hideTooltip();
            });

            this._waypointsGroup.appendChild(circle);
            this._waypointCircles.push(circle);
        }
    }

    private showTooltip(point: RoutePoint, index: number, total: number): void {
        if (!this._map || !this._tooltipEl) return;

        const pixels = this._map.positionsToPixels([[point.lng, point.lat]]);
        if (!pixels || pixels.length === 0) return;

        const px = pixels[0];
        if (!Number.isFinite(px[0]) || !Number.isFinite(px[1])) return;

        let label = `Waypoint #${index}`;
        let badgeColor = "#004b8b";
        if (index === 0) {
            label = "Inicio de ruta";
            badgeColor = "#16834b";
        } else if (index === total - 1) {
            label = `Fin de ruta (#${index})`;
            badgeColor = "#c43131";
        }

        const timeStr = this.formatTimestamp(point.timestamp);
        const coordsStr = `${point.lat.toFixed(5)}, ${point.lng.toFixed(5)}`;

        this._tooltipEl.innerHTML =
            `<div style="font-weight: 600; color: ${badgeColor}; margin-bottom: 3px;">${label}</div>` +
            `<div><strong>Hora:</strong> ${timeStr}</div>` +
            `<div><strong>Geoposición:</strong> ${coordsStr}</div>`;

        this._tooltipEl.style.left = `${px[0]}px`;
        this._tooltipEl.style.top = `${px[1]}px`;
        this._tooltipEl.style.transform = px[1] < 85
            ? "translate(-50%, 16px)"
            : "translate(-50%, calc(-100% - 16px))";
        this._tooltipEl.style.display = "block";
    }

    private hideTooltip(): void {
        if (this._tooltipEl) {
            this._tooltipEl.style.display = "none";
        }
    }

    private formatTimestamp(isoString?: string): string {
        if (!isoString) return "Sin hora";
        const d = new Date(isoString);
        if (isNaN(d.getTime())) return "Sin hora";
        return d.toLocaleTimeString("es-ES", {
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit"
        });
    }

    private drawRoute(): void {
        if (!this._map || !this._routeSvg || !this._routeHalo || !this._routeLine || !this._startMarker || !this._endMarker) return;

        if (this._routePositions.length === 0) {
            this._routeHalo.setAttribute("points", "");
            this._routeLine.setAttribute("points", "");
            this._startMarker.setAttribute("display", "none");
            this._endMarker.setAttribute("display", "none");
            for (const c of this._waypointCircles) {
                c.setAttribute("display", "none");
            }
            return;
        }

        const pixels = this._map.positionsToPixels(this._routePositions);
        if (pixels.some(pixel => !Number.isFinite(pixel[0]) || !Number.isFinite(pixel[1]))) return;

        if (this._routePositions.length === 1) {
            this._routeHalo.setAttribute("points", "");
            this._routeLine.setAttribute("points", "");
            this._startMarker.setAttribute("transform", `translate(${pixels[0][0]} ${pixels[0][1]})`);
            this._startMarker.removeAttribute("display");
            this._endMarker.setAttribute("display", "none");
            return;
        }

        const path = pixels.map(pixel => `${pixel[0]},${pixel[1]}`).join(" ");
        this._routeHalo.setAttribute("points", path);
        this._routeLine.setAttribute("points", path);

        for (let i = 0; i < this._waypointCircles.length; i++) {
            const p = pixels[i + 1];
            const circle = this._waypointCircles[i];
            if (p) {
                circle.setAttribute("cx", String(p[0]));
                circle.setAttribute("cy", String(p[1]));
                circle.removeAttribute("display");
            } else {
                circle.setAttribute("display", "none");
            }
        }

        const lastPixel = pixels[pixels.length - 1];
        this._startMarker.setAttribute("transform", `translate(${pixels[0][0]} ${pixels[0][1]})`);
        this._endMarker.setAttribute("transform", `translate(${lastPixel[0]} ${lastPixel[1]})`);
        this._startMarker.removeAttribute("display");
        this._endMarker.removeAttribute("display");
    }

    private changeMapStyle(option?: IDropdownOption): void {
        if (!option) return;
        this._selectedMapStyle = option.key as string;
        this.applyMapStyleLayers();
        this.renderHeader();
    }

    private async loadResources(): Promise<void> {
        try {
            const queryOptions = "?$select=name,_userid_value&$filter=statecode eq 0 and _userid_value ne null&$orderby=name asc";
            const result = await this._webAPI.retrieveMultipleRecords("bookableresource", queryOptions);

            this._resources = result.entities.map((entity: ComponentFramework.WebApi.Entity) => ({
                key: entity._userid_value as string,
                text: entity.name as string
            }));

            if (this._resources.length > 0 && !this._selectedResource) {
                this._selectedResource = this._resources[0].key as string;
            }

            this.setStatus(`${this._resources.length} recursos cargados.`);
        } catch (error) {
            console.error("[SecuRutas] Error cargando recursos:", error);
            this.setStatus("⚠ Error cargando recursos: " + this.errMsg(error));
        }
    }

    private async fetchAll(entityName: string, firstOptions: string): Promise<ComponentFramework.WebApi.Entity[]> {
        const all: ComponentFramework.WebApi.Entity[] = [];
        let options = firstOptions;

        for (let page = 0; page < this._maxPages; page++) {
            const result = await this._webAPI.retrieveMultipleRecords(entityName, options, this._pageSize);
            all.push(...result.entities);

            if (!result.nextLink) break;
            const idx = result.nextLink.indexOf("?");
            if (idx < 0) break;
            options = result.nextLink.substring(idx);
        }
        return all;
    }

    private async fetchTrackingData(): Promise<void> {
        if (!this._selectedResource || !this._selectedDate) {
            this.setStatus("⚠ Selecciona una fecha y un recurso.");
            return;
        }

        this._loading = true;
        this.setStatus("Buscando…");

        try {
            const startOfDay = new Date(this._selectedDate);
            startOfDay.setHours(0, 0, 0, 0);
            const endOfDay = new Date(this._selectedDate);
            endOfDay.setHours(23, 59, 59, 999);

            const startStr = startOfDay.toISOString();
            const endStr = endOfDay.toISOString();

            const options = "?$select=msdyn_latitude,msdyn_longitude,msdyn_date" +
                "&$filter=_msdyn_userid_value eq " + this._selectedResource +
                " and msdyn_date ge " + startStr + " and msdyn_date le " + endStr +
                "&$orderby=msdyn_date asc";

            const entities = await this.fetchAll("msdyn_geolocationtracking", options);

            const points: RoutePoint[] = [];
            for (const entity of entities) {
                if (entity.msdyn_latitude == null || entity.msdyn_longitude == null) continue;

                const lat = parseFloat(entity.msdyn_latitude.toString().replace(",", "."));
                const lng = parseFloat(entity.msdyn_longitude.toString().replace(",", "."));

                if (isNaN(lat) || isNaN(lng) || (lat === 0 && lng === 0)) continue;

                const last = points[points.length - 1];
                if (!last || last.lng !== lng || last.lat !== lat) {
                    points.push({
                        lng,
                        lat,
                        timestamp: entity.msdyn_date as string | undefined
                    });
                }
            }

            if (entities.length === 0) {
                await this.explainEmptyResult();
            } else if (points.length === 0) {
                this.setStatus(`${entities.length} registros, pero ninguno con coordenadas válidas.`);
            } else if (points.length === 1) {
                this.setStatus(`${entities.length} registros, pero todos en la misma posición.`);
            } else {
                this.setStatus(`${entities.length} registros → ${points.length} puntos en la ruta.`);
            }

            if (!this._isMapReady) {
                this._pendingPoints = points;
            } else {
                this.updateMap(points);
            }
        } catch (error) {
            console.error("[SecuRutas] Error obteniendo tracking:", error);
            this.setStatus("⚠ Error en la consulta: " + this.errMsg(error));
            this.updateMap([]);
        } finally {
            this._loading = false;
            this.renderHeader();
        }
    }

    private async explainEmptyResult(): Promise<void> {
        try {
            const queryOptions = "?$select=msdyn_date&$filter=_msdyn_userid_value eq " + this._selectedResource + "&$orderby=msdyn_date desc&$top=1";
            const res = await this._webAPI.retrieveMultipleRecords("msdyn_geolocationtracking", queryOptions);

            if (res.entities.length === 0) {
                this.setStatus("Sin registros para este recurso en ninguna fecha (¿es el usuario correcto?).");
            } else {
                const d = new Date(res.entities[0].msdyn_date as string);
                this.setStatus(`Sin registros ese día. Último registro del recurso: ${d.toLocaleString()}.`);
            }
        } catch (e) {
            this.setStatus("Sin registros para ese día.");
        }
    }

    private setStatus(message: string): void {
        this._status = message;
        console.log("[SecuRutas]", message);
        this.renderHeader();
    }

    private errMsg(error: unknown): string {
        if (error && typeof error === "object" && "message" in error) {
            return String((error as { message: unknown }).message);
        }
        return String(error);
    }

    private renderHeader(): void {
        const control = React.createElement(
            "div",
            { className: "modern-header-card" },
            React.createElement(
                "div",
                { className: "modern-header-title" },
                React.createElement("h2", null, "Seguimiento de Rutas")
            ),
            React.createElement(
                "div",
                { className: "modern-header-controls" },
                React.createElement(DatePicker, {
                    label: "Fecha",
                    value: this._selectedDate,
                    onSelectDate: (date: Date | null | undefined) => {
                        if (date) {
                            this._selectedDate = date;
                            this.renderHeader();
                        }
                    },
                    styles: { root: { width: 160 } }
                }),
                React.createElement(ComboBox, {
                    label: "Recurso Técnico",
                    selectedKey: this._selectedResource,
                    allowFreeform: true,
                    autoComplete: "on",
                    onChange: (event: React.FormEvent<IComboBox>, option?: IComboBoxOption) => {
                        if (option) {
                            this._selectedResource = option.key as string;
                            this.renderHeader();
                        }
                    },
                    options: this._resources,
                    styles: { root: { width: 320 } }
                }),
                React.createElement(Dropdown, {
                    label: "Tipo de mapa",
                    selectedKey: this._selectedMapStyle,
                    options: [
                        { key: "road", text: "Normal (Azure)" },
                        { key: "osm", text: "Callejero (OSM)" },
                        { key: "satellite_road_labels", text: "Satélite" }
                    ],
                    onChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
                        this.changeMapStyle(option);
                    },
                    styles: { root: { width: 180 } }
                }),
                React.createElement(IconButton, {
                    iconProps: { iconName: "Search" },
                    title: "Buscar",
                    disabled: this._loading,
                    onClick: () => { void this.fetchTrackingData(); },
                    className: "modern-search-button"
                })
            ),
            this._status
                ? React.createElement(
                    "div",
                    { style: { marginTop: 6, fontSize: 12, color: "#605e5c" } },
                    this._status
                )
                : null,
            this._mapError
                ? React.createElement(
                    "div",
                    { style: { marginTop: 2, fontSize: 12, color: "#a4262c" } },
                    this._mapError
                )
                : null
        );

        if (!this._root) {
            this._root = createRoot(this._headerContainer);
        }
        this._root.render(control);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this.safeResize();
    }

    public getOutputs(): IOutputs { return {}; }

    public destroy(): void {
        if (this._resizeObserver) this._resizeObserver.disconnect();
        if (this._root) this._root.unmount();
        if (this._map) this._map.dispose();
        if (this._routeSvg) this._routeSvg.remove();
        if (this._tooltipEl) this._tooltipEl.remove();
    }
}