import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as atlas from "azure-maps-control";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { ComboBox, IComboBoxOption, IComboBox, DatePicker, IconButton, IIconProps, initializeIcons } from '@fluentui/react';

// Inicializar iconos de Fluent UI para la lupa
initializeIcons(undefined, { disableWarnings: true });

interface GeoLocationPoint {
    id: string;
    lat: number;
    lng: number;
    timestamp: Date;
}

interface MarkerProperties {
    icon: string;
    title: string;
}

export class SecuRutas implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _mapContainer: HTMLDivElement;
    private _headerContainer: HTMLDivElement;
    private _versionContainer: HTMLDivElement;
    private _map: atlas.Map;
    private _dataSource: atlas.source.DataSource;
    
    private _selectedDate: Date = new Date();
    private _selectedResource: string | undefined;
    private _resources: IComboBoxOption[] = [];
    
    private _webAPI: ComponentFramework.WebApi;
    private _root: Root;
    private readonly _controlVersion: string = "v1.0.9";

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._container = container;
        this._webAPI = context.webAPI;

        const mainDiv = document.createElement("div");
        mainDiv.className = "secu-rutas-container";

        this._headerContainer = document.createElement("div");
        this._headerContainer.className = "secu-rutas-header";

        this._mapContainer = document.createElement("div");
        this._mapContainer.className = "secu-rutas-map";

        this._versionContainer = document.createElement("div");
        this._versionContainer.className = "secu-rutas-version";
        this._versionContainer.innerText = this._controlVersion;

        mainDiv.appendChild(this._headerContainer);
        mainDiv.appendChild(this._mapContainer);
        mainDiv.appendChild(this._versionContainer);
        this._container.appendChild(mainDiv);

        this.initMap();
        this.loadResources();
    }

    private initMap(): void {
        const azureMapsKey = this._context.parameters.azureMapsKey.raw;
        const azureMapsUrl = this._context.parameters.azureMapsUrl.raw;

        if (!azureMapsKey) {
            this._mapContainer.innerHTML = "<div style='padding: 20px; color: red;'>Falta la clave de Azure Maps en la configuración.</div>";
            return;
        }

        const mapOptions: atlas.ServiceOptions & atlas.StyleOptions & atlas.UserInteractionOptions & (atlas.CameraOptions | atlas.CameraBoundsOptions) = {
            authOptions: {
                authType: atlas.AuthenticationType.subscriptionKey,
                subscriptionKey: azureMapsKey
            },
            view: 'Auto',
            style: 'road'
        };

        this._map = new atlas.Map(this._mapContainer, mapOptions);

        this._map.events.add('ready', () => {
            this._dataSource = new atlas.source.DataSource();
            this._map.sources.add(this._dataSource);

            this._map.layers.add(new atlas.layer.LineLayer(this._dataSource, undefined, {
                strokeColor: '#004b8b', // Azul oscuro acorde al nuevo diseño
                strokeWidth: 4
            }));

            this._map.layers.add(new atlas.layer.SymbolLayer(this._dataSource, undefined, {
                iconOptions: {
                    image: ['get', 'icon'],
                    size: 1.2
                },
                textOptions: {
                    textField: ['get', 'title'],
                    offset: [0, 1.2],
                    color: '#004b8b'
                }
            }));
            
            this._map.setCamera({
                center: [-3.7038, 40.4168],
                zoom: 10
            });
            
            this.updateMap();
        });
        
        setTimeout(() => {
            if (this._map) this._map.events.invoke('resize', null);
        }, 500);
    }

    private async loadResources(): Promise<void> {
        try {
            const result = await this._webAPI.retrieveMultipleRecords("bookableresource", "?$select=bookableresourceid,name&$filter=statecode eq 0");
            
            this._resources = result.entities.map((entity: ComponentFramework.WebApi.Entity) => ({
                key: entity.bookableresourceid,
                text: entity.name
            }));

            this.renderHeader();
            
            if (this._resources.length > 0 && !this._selectedResource) {
                this._selectedResource = this._resources[0].key as string;
                this.renderHeader(); 
            }

        } catch (error) {
            console.error("[SecuRutas] Error cargando recursos:", error);
        }
    }

    private renderHeader(): void {
        const searchIcon: IIconProps = { iconName: 'Search' };

        const control = React.createElement(
            'div',
            { className: 'modern-header-card' },
            React.createElement(
                'div',
                { className: 'modern-header-title' },
                React.createElement('h2', null, 'Seguimiento de Rutas')
            ),
            React.createElement(
                'div',
                { className: 'modern-header-controls' },
                React.createElement(DatePicker, {
                    label: "Fecha",
                    value: this._selectedDate,
                    onSelectDate: (date: Date | null | undefined) => {
                        if (date) {
                            this._selectedDate = date;
                            // Solo repinta, NO busca
                            this.renderHeader(); 
                        }
                    },
                    styles: { root: { width: 160 } }
                }),
                React.createElement(ComboBox, {
                    label: "Recurso Técnico",
                    selectedKey: this._selectedResource,
                    allowFreeform: true,
                    autoComplete: 'on',
                    onChange: (event: React.FormEvent<IComboBox>, option?: IComboBoxOption) => {
                        if (option) {
                            this._selectedResource = option.key as string;
                            // Solo repinta, NO busca
                            this.renderHeader(); 
                        }
                    },
                    placeholder: "Buscar recurso...",
                    options: this._resources,
                    styles: { root: { width: 320 } }
                }),
                React.createElement(IconButton, {
                    iconProps: searchIcon,
                    title: "Buscar rutas",
                    ariaLabel: "Buscar rutas",
                    onClick: () => {
                        this.fetchTrackingData();
                    },
                    className: 'modern-search-button'
                })
            )
        );

        if (!this._root) {
            this._root = createRoot(this._headerContainer);
        }
        
        this._root.render(control);
    }

    private async fetchTrackingData(): Promise<void> {
        if (!this._selectedResource || !this._selectedDate) {
            return;
        }

        try {
            const startOfDay = new Date(this._selectedDate);
            startOfDay.setHours(0, 0, 0, 0);
            const endOfDay = new Date(this._selectedDate);
            endOfDay.setHours(23, 59, 59, 999);

            const startStr = startOfDay.toISOString();
            const endStr = endOfDay.toISOString();

            const query = `?$select=msdyn_latitude,msdyn_longitude,createdon&$filter=_msdyn_resource_value eq ${this._selectedResource} and createdon ge ${startStr} and createdon le ${endStr}&$orderby=createdon asc`;
            
            console.log(`[SecuRutas] Ejecutando consulta WebAPI:`, query);

            const result = await this._webAPI.retrieveMultipleRecords("msdyn_geolocationtracking", query);
            
            if (!result || !result.entities || result.entities.length === 0) {
                 console.log("[SecuRutas] No hay puntos para mostrar. Limpiando mapa.");
                 this.updateMap([]);
                 return;
            }
            
            const points: GeoLocationPoint[] = result.entities.map((entity: ComponentFramework.WebApi.Entity) => ({
                id: entity.msdyn_geolocationtrackingid,
                lat: parseFloat(entity.msdyn_latitude),
                lng: parseFloat(entity.msdyn_longitude),
                timestamp: new Date(entity.createdon)
            }));

            this.updateMap(points);

        } catch (error) {
            console.error("[SecuRutas] FATAL ERROR consultando tracking:", error);
            if (error instanceof Error) {
                 console.error("[SecuRutas] Detalle:", error.message);
            }
            this.updateMap([]);
        }
    }

    private updateMap(points: GeoLocationPoint[] = []): void {
        if (!this._map || !this._dataSource) return;

        this._dataSource.clear();

        if (points.length === 0) {
             this._map.setCamera({
                center: [-3.7038, 40.4168],
                zoom: 10
            });
            return;
        }

        const coordinates = points.map(p => [p.lng, p.lat]);
        const line = new atlas.data.LineString(coordinates);
        this._dataSource.add(line);

        const markers: atlas.data.Feature<atlas.data.Point, MarkerProperties>[] = [];

        points.forEach((point, index) => {
            let icon = 'marker-blue';
            let title = '';

            if (index === 0) {
                icon = 'marker-green';
                title = 'Inicio';
            } else if (index === points.length - 1) {
                icon = 'marker-red';
                title = 'Fin';
            }

            if (index === 0 || index === points.length - 1) {
                 markers.push(new atlas.data.Feature(new atlas.data.Point([point.lng, point.lat]), {
                    icon: icon,
                    title: title
                }));
            } else {
                markers.push(new atlas.data.Feature(new atlas.data.Point([point.lng, point.lat]), {
                    icon: 'pin-round-blue',
                    title: '' 
                }));
            }
        });

        this._dataSource.add(markers);

        const bounds = atlas.data.BoundingBox.fromData(line);
        this._map.setCamera({
            bounds: bounds,
            padding: 50 
        });
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        if (this._map) {
             setTimeout(() => {
                this._map.events.invoke('resize', null);
            }, 100);
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        if (this._root) {
            this._root.unmount();
        }
        if (this._map) {
            this._map.dispose();
        }
    }
}