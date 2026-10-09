import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as atlas from "azure-maps-control";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { 
    ComboBox, IComboBoxOption, IComboBox, DatePicker, IconButton, IIconProps, 
    initializeIcons, Dialog, DialogType, DialogFooter, PrimaryButton 
} from '@fluentui/react';

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
    tooltipInfo: string;
}

export class SecuRutas implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _mapContainer: HTMLDivElement;
    private _headerContainer: HTMLDivElement;
    private _versionContainer: HTMLDivElement;
    
    private _map: atlas.Map;
    // SOLUCIÓN: Separamos las fuentes de datos para evitar conflictos de geometría
    private _lineSource: atlas.source.DataSource;
    private _pointSource: atlas.source.DataSource;
    private _symbolLayer: atlas.layer.SymbolLayer;
    private _hoverPopup: atlas.Popup;
    
    private _selectedDate: Date = new Date();
    private _selectedResource: string | undefined;
    private _resources: IComboBoxOption[] = [];
    
    private _showDialog = false;
    private _dialogTitle = "";
    private _dialogContent: React.ReactNode = null;

    private _webAPI: ComponentFramework.WebApi;
    private _root: Root;
    private readonly _controlVersion = "v1.1.7";

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
            // Inicializar fuentes independientes
            this._lineSource = new atlas.source.DataSource();
            this._pointSource = new atlas.source.DataSource();
            this._map.sources.add(this._lineSource);
            this._map.sources.add(this._pointSource);

            // Capa exclusiva para la ruta (ya no requiere filtros)
            this._map.layers.add(new atlas.layer.LineLayer(this._lineSource, undefined, {
                strokeColor: '#004b8b',
                strokeWidth: 4
            }));

            // Capa exclusiva para los waypoints (ya no requiere filtros)
            this._symbolLayer = new atlas.layer.SymbolLayer(this._pointSource, undefined, {
                iconOptions: {
                    image: ['get', 'icon'], // Leerá directamente la propiedad 'icon'
                    allowOverlap: true, 
                    ignorePlacement: true
                }
            });
            this._map.layers.add(this._symbolLayer);
            
            // Inicializar el Popup para el Tooltip
            this._hoverPopup = new atlas.Popup({
                pixelOffset: [0, -18],
                closeButton: false,
                fillColor: 'rgba(255,255,255,0.95)'
            });

            // Eventos del Tooltip atados a la capa de puntos
            this._map.events.add('mouseover', this._symbolLayer, (e) => {
                if (e.shapes && e.shapes.length > 0) {
                    const shape = e.shapes[0];
                    if (shape instanceof atlas.Shape) {
                        const properties = shape.getProperties();
                        const position = shape.getCoordinates() as atlas.data.Position;
                        
                        this._hoverPopup.setOptions({
                            content: `<div style="padding: 10px; font-family: 'Segoe UI', sans-serif; font-size: 12px; color: #333; min-width: 120px;">
                                        <strong style="color: #004b8b; font-size: 13px;">${properties.title}</strong><br/>
                                        <div style="margin-top: 5px;">${properties.tooltipInfo}</div>
                                      </div>`,
                            position: position
                        });
                        this._hoverPopup.open(this._map);
                    }
                }
            });

            this._map.events.add('mouseleave', this._symbolLayer, () => {
                this._hoverPopup.close();
            });
            
            this._map.setCamera({
                center: [-3.7038, 40.4168],
                zoom: 10
            });
            
        });
        
        setTimeout(() => {
            if (this._map) this._map.events.invoke('resize', null);
        }, 500);
    }

    private async loadResources(): Promise<void> {
        try {
            const result = await this._webAPI.retrieveMultipleRecords("bookableresource", "?$select=name,_userid_value&$filter=statecode eq 0 and _userid_value ne null");
            
            this._resources = result.entities.map((entity: ComponentFramework.WebApi.Entity) => ({
                key: entity._userid_value, 
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

    private closeDialog = (): void => {
        this._showDialog = false;
        this.renderHeader();
    }

    private renderHeader(): void {
        const searchIcon: IIconProps = { iconName: 'Search' };

        const headerCard = React.createElement(
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

        const dialog = React.createElement(Dialog, {
            hidden: !this._showDialog,
            onDismiss: this.closeDialog,
            dialogContentProps: {
                type: DialogType.largeHeader,
                title: this._dialogTitle
            }
        },
            this._dialogContent,
            React.createElement(DialogFooter, null,
                React.createElement(PrimaryButton, { onClick: this.closeDialog, text: "Aceptar" })
            )
        );

        const control = React.createElement(React.Fragment, null, headerCard, dialog);

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

            const query = `?$select=msdyn_latitude,msdyn_longitude,createdon&$filter=_msdyn_userid_value eq ${this._selectedResource} and createdon ge ${startStr} and createdon le ${endStr}&$orderby=createdon asc`;
            
            const result = await this._webAPI.retrieveMultipleRecords("msdyn_geolocationtracking", query);
            
            if (!result || !result.entities || result.entities.length === 0) {
                 this.renderHeader();
                 this.updateMap([]);
                 return;
            }
            
            const points: GeoLocationPoint[] = [];
            
            result.entities.forEach((entity: ComponentFramework.WebApi.Entity) => {
                if (entity.msdyn_latitude != null && entity.msdyn_longitude != null) {
                    const lat = parseFloat(entity.msdyn_latitude);
                    const lng = parseFloat(entity.msdyn_longitude);
                    
                    if (!isNaN(lat) && !isNaN(lng) && (lat !== 0 || lng !== 0)) {
                        points.push({
                            id: entity.msdyn_geolocationtrackingid,
                            lat: lat,
                            lng: lng,
                            timestamp: new Date(entity.createdon)
                        });
                    }
                }
            });

            this.updateMap(points);

        } catch (error) {
            console.error("[SecuRutas] FATAL ERROR consultando tracking:", error);
            const errorMsg = error instanceof Error ? error.message : JSON.stringify(error);
            
            this._dialogTitle = "Error en la Búsqueda (Dataverse)";
            this._dialogContent = React.createElement('div', null,
                React.createElement('p', null, "Ha fallado la consulta a la tabla msdyn_geolocationtracking."),
                React.createElement('pre', { style: { backgroundColor: '#ffe5e5', padding: '10px', borderRadius: '4px', overflowX: 'auto', fontSize: '12px' } }, errorMsg)
            );
            this._showDialog = true;
            this.renderHeader();

            this.updateMap([]);
        }
    }

    private updateMap(points: GeoLocationPoint[] = []): void {
        if (!this._map || !this._lineSource || !this._pointSource) return;

        // Limpiamos fuentes independientes
        this._lineSource.clear();
        this._pointSource.clear();
        if (this._hoverPopup) this._hoverPopup.close();

        if (points.length === 0) {
             this._map.setCamera({
                center: [-3.7038, 40.4168],
                zoom: 10
            });
            return;
        }

        // 1. Pintar la línea en su propio DataSource
        if (points.length > 1) {
            const coordinates: atlas.data.Position[] = points.map(point => [point.lng, point.lat]);
            const route = new atlas.data.Feature(new atlas.data.LineString(coordinates));
            this._lineSource.add(route);
        }

        // 2. Pintar los marcadores en su propio DataSource
        const markers: atlas.data.Feature<atlas.data.Point, MarkerProperties>[] = [];

        points.forEach((point, index) => {
            let icon = 'pin-round-blue'; 
            let title = 'Waypoint';

            if (index === 0) {
                icon = 'marker-blue';
                title = 'Inicio';
            } else if (index === points.length - 1) {
                icon = 'marker-red';
                title = 'Fin';
            }

            const timeString = point.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            
            markers.push(new atlas.data.Feature(new atlas.data.Point([point.lng, point.lat]), {
                icon: icon,
                title: title,
                tooltipInfo: `<strong>Hora:</strong> ${timeString}<br/><strong>Lat:</strong> ${point.lat.toFixed(5)}<br/><strong>Lng:</strong> ${point.lng.toFixed(5)}`
            }));
        });

        this._pointSource.add(markers);

        // 3. Ajustar la cámara para que englobe todos los puntos
        const lats = points.map(p => p.lat);
        const lngs = points.map(p => p.lng);
        
        const bounds = atlas.data.BoundingBox.fromEdges(
            Math.min(...lngs),
            Math.min(...lats),
            Math.max(...lngs),
            Math.max(...lats)
        );

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