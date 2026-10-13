import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as atlas from "azure-maps-control";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { ComboBox, IComboBoxOption, IComboBox, DatePicker, IconButton, initializeIcons } from '@fluentui/react';

initializeIcons(undefined, { disableWarnings: true });

export class SecuRutas implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _mapContainer: HTMLDivElement;
    private _headerContainer: HTMLDivElement;

    private _map: atlas.Map | undefined;
    private _dataSource: atlas.source.DataSource | undefined;
    private _routeSvg: SVGSVGElement | undefined;
    private _routeHalo: SVGPolylineElement | undefined;
    private _routeLine: SVGPolylineElement | undefined;
    private _routePositions: atlas.data.Position[] = [];
    private _isMapReady = false;
    private _resizeObserver: ResizeObserver | undefined;
    private _pendingCoords: number[][] | undefined;

    private _selectedDate: Date = new Date();
    private _selectedResource: string | undefined;
    private _resources: IComboBoxOption[] = [];

    private _status = "";
    private _mapError = "";
    private _loading = false;

    private _webAPI: ComponentFramework.WebApi;
    private _root: Root;
    private readonly _controlVersion = "v1.1.30";

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
            center: [-3.7038, 40.4168],
            zoom: 10,
            style: "road"
        });
        this._map = map;

        map.events.add("error", (e: unknown) => {
            console.error("[SecuRutas] Error del mapa:", e);
            const inner = (e as { error?: { message?: string } })?.error;
            this._mapError = "⚠ Azure Maps: " + (inner?.message ?? this.errMsg(e));
            this.renderHeader();
        });

        map.events.add("ready", () => {
            const ds = new atlas.source.DataSource();
            map.sources.add(ds);
            this._dataSource = ds;

            map.layers.add(
                new atlas.layer.SymbolLayer(ds, undefined, {
                    filter: ["==", ["geometry-type"], "Point"],
                    iconOptions: {
                        image: ["coalesce", ["get", "icon"], "marker-blue"],
                        allowOverlap: true,
                        ignorePlacement: true
                    },
                    textOptions: {
                        textField: ["get", "title"],
                        offset: [0, 1.2]
                    }
                })
            );

            this.createRouteOverlay();
            map.events.add("render", () => this.drawRoute());
            this._isMapReady = true;
            this.safeResize();

            if (this._pendingCoords) {
                const pending = this._pendingCoords;
                this._pendingCoords = undefined;
                this.updateMap(pending);
            }
        });

        if (typeof ResizeObserver !== "undefined") {
            this._resizeObserver = new ResizeObserver(() => this.safeResize());
            this._resizeObserver.observe(this._mapContainer);
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

    private updateMap(coords: number[][]): void {
        if (!this._isMapReady || !this._dataSource || !this._map) {
            this._pendingCoords = coords;
            return;
        }

        const ds = this._dataSource;
        ds.clear();
        this._routePositions = coords.map(coord => [coord[0], coord[1]]);

        if (coords.length === 0) {
            this.drawRoute();
            return;
        }

        const features: atlas.data.Feature<atlas.data.Point, { icon: string; title: string }>[] = [];

        coords.forEach((c, index) => {
            let icon = "pin-round-blue";
            let title = "";

            if (index === 0) {
                icon = "marker-blue";
                title = "Inicio";
            } else if (index === coords.length - 1) {
                icon = "marker-red";
                title = "Fin";
            }

            features.push(new atlas.data.Feature(
                new atlas.data.Point([c[0], c[1]]),
                { icon, title }
            ));
        });

        ds.add(new atlas.data.FeatureCollection(features));
        this.drawRoute();

        let minLat = Infinity, maxLat = -Infinity, minLng = Infinity, maxLng = -Infinity;
        for (const c of coords) {
            if (c[0] < minLng) minLng = c[0];
            if (c[0] > maxLng) maxLng = c[0];
            if (c[1] < minLat) minLat = c[1];
            if (c[1] > maxLat) maxLat = c[1];
        }

        this.safeResize();

        if (Math.abs(maxLat - minLat) > 0.0001 || Math.abs(maxLng - minLng) > 0.0001) {
            this._map.setCamera({
                bounds: atlas.data.BoundingBox.fromEdges(minLng, minLat, maxLng, maxLat),
                padding: 50
            });
        } else {
            this._map.setCamera({ center: [coords[0][0], coords[0][1]], zoom: 15 });
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
        svg.style.zIndex = "1";

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

        svg.appendChild(halo);
        svg.appendChild(line);
        this._mapContainer.appendChild(svg);
        this._routeSvg = svg;
        this._routeHalo = halo;
        this._routeLine = line;
    }

    private drawRoute(): void {
        if (!this._map || !this._routeSvg || !this._routeHalo || !this._routeLine) return;
        if (this._routePositions.length < 2) {
            this._routeHalo.setAttribute("points", "");
            this._routeLine.setAttribute("points", "");
            return;
        }

        const pixels = this._map.positionsToPixels(this._routePositions);
        if (pixels.some(pixel => !Number.isFinite(pixel[0]) || !Number.isFinite(pixel[1]))) return;

        const path = pixels.map(pixel => `${pixel[0]},${pixel[1]}`).join(" ");
        this._routeHalo.setAttribute("points", path);
        this._routeLine.setAttribute("points", path);
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

            const coords: number[][] = [];
            for (const entity of entities) {
                if (entity.msdyn_latitude == null || entity.msdyn_longitude == null) continue;

                const lat = parseFloat(entity.msdyn_latitude.toString().replace(",", "."));
                const lng = parseFloat(entity.msdyn_longitude.toString().replace(",", "."));

                if (isNaN(lat) || isNaN(lng) || (lat === 0 && lng === 0)) continue;

                const last = coords[coords.length - 1];
                if (!last || last[0] !== lng || last[1] !== lat) {
                    coords.push([lng, lat]);
                }
            }

            if (entities.length === 0) {
                await this.explainEmptyResult();
            } else if (coords.length === 0) {
                this.setStatus(`${entities.length} registros, pero ninguno con coordenadas válidas.`);
            } else if (coords.length === 1) {
                this.setStatus(`${entities.length} registros, pero todos en la misma posición: no hay línea que dibujar.`);
            } else {
                this.setStatus(`${entities.length} registros → ${coords.length} puntos en la ruta.`);
            }

            if (!this._isMapReady) {
                this._pendingCoords = coords;
            } else {
                this.updateMap(coords);
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
    }
}