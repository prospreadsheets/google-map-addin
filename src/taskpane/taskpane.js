/* global Office, Excel, google, html2canvas */

let map         = null;
let markers     = [];
let apiKey      = null;
let mapId       = null;
let mapData     = null;
let lastZoom    = null;
let lastCenterLat = null;
let lastCenterLng = null;

Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        setStatus("Ready — click Load Map");
        document.getElementById("btnLoadMap").disabled = false;
    } else {
        setStatus("Error: Not running in Excel");
    }
});

function setStatus(msg) {
    document.getElementById("status").textContent = msg;
}

function loadMap() {
    setStatus("Reading data from Excel...");
    document.getElementById("btnLoadMap").disabled  = true;
    document.getElementById("btnCapture").disabled  = true;
    document.getElementById("btnReset").disabled    = true;

    Excel.run(function(context) {
        var apiKeyRange  = context.workbook.names.getItem("GMAP_API_KEY").getRange();
        var mapIdRange   = context.workbook.names.getItem("GMAP_MAP_ID").getRange();
        var dataRange    = context.workbook.names.getItem("GMAP_DATA").getRange();
        var zoomRange    = context.workbook.names.getItem("GMAP_LAST_ZOOM").getRange();
        var latRange     = context.workbook.names.getItem("GMAP_LAST_LAT").getRange();
        var lngRange     = context.workbook.names.getItem("GMAP_LAST_LNG").getRange();

        apiKeyRange.load("values");
        mapIdRange.load("values");
        dataRange.load("values");
        zoomRange.load("values");
        latRange.load("values");
        lngRange.load("values");

        return context.sync().then(function() {
            apiKey = apiKeyRange.values[0][0];
            mapId  = mapIdRange.values[0][0] || "";
            var raw = dataRange.values[0][0];

            var savedZoom = zoomRange.values[0][0];
            var savedLat  = latRange.values[0][0];
            var savedLng  = lngRange.values[0][0];

            if (savedZoom && savedZoom !== "") {
                lastZoom      = parseInt(savedZoom);
                lastCenterLat = parseFloat(savedLat);
                lastCenterLng = parseFloat(savedLng);
            }

            if (!apiKey || !raw) {
                setStatus("Error: No data. Run the macro first.");
                document.getElementById("btnLoadMap").disabled = false;
                return;
            }

            try {
                mapData = JSON.parse(raw);
            } catch(e) {
                setStatus("Error: Could not parse map data.");
                document.getElementById("btnLoadMap").disabled = false;
                return;
            }

            loadGoogleMapsAPI(apiKey, mapId);
        });
    }).catch(function(err) {
        setStatus("Error reading Excel: " + err.message);
        document.getElementById("btnLoadMap").disabled = false;
    });
}

function loadGoogleMapsAPI(key, mid) {
    const existing = document.getElementById("gmaps-script");
    if (existing) existing.remove();

    setStatus("Loading Google Maps...");
    window._gmapMapId = mid;

    const script    = document.createElement("script");
    script.id       = "gmaps-script";
    script.src      = "https://maps.googleapis.com/maps/api/js?key=" + key + "&callback=initMap";
    script.async    = true;
    script.defer    = true;
    script.onerror  = function() {
        setStatus("Error: Could not load Google Maps. Check API key.");
        document.getElementById("btnLoadMap").disabled = false;
    };
    document.head.appendChild(script);
}

function initMap() {
    setStatus("Plotting markers...");

    markers.forEach(function(m) { m.setMap(null); });
    markers = [];

    const mapDiv = document.getElementById("map");

    const mapOptions = {
        zoom:               lastZoom || 12,
        center:             { lat: lastCenterLat || 0, lng: lastCenterLng || 0 },
        mapTypeId:          "roadmap",
        fullscreenControl:  false,
        streetViewControl:  false,
        mapTypeControl:     true,
        zoomControl:        true
    };

    if (window._gmapMapId && window._gmapMapId !== "") {
        mapOptions.mapId = window._gmapMapId;
    }

    map = new google.maps.Map(mapDiv, mapOptions);

    const bounds          = new google.maps.LatLngBounds();
    let centerSet         = false;
    let centerLatLng      = null;
    let hasVisibleMarkers = false;

    mapData.forEach(function(row) {
        const colDVal  = (row.flag || "").toLowerCase().trim();
        const hideAll  = colDVal.indexOf("hide all")  > -1;
        const hidePin  = hideAll ||
                         colDVal === "yes" ||
                         colDVal === "y"   ||
                         colDVal.indexOf("hide") > -1;
        const isCenter = colDVal.indexOf("center") > -1;

        // Capture first center flagged row
        if (isCenter && !centerSet) {
            centerLatLng = { lat: row.lat, lng: row.lng };
            centerSet    = true;
        }

        // Skip hidden pins
        if (hidePin) return;

        const position  = new google.maps.LatLng(row.lat, row.lng);
        const fillColor = row.fillColor || "#FF0000";
        const fontColor = row.fontColor || "#FFFFFF";
        const labelText = String(row.label);
        const charCount = labelText.length;
        const markerW   = Math.max(32, charCount * 10 + 12);
        const markerH   = 36;
        const fontSize  = charCount <= 2 ? 13 : charCount <= 4 ? 11 : 9;

        const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${markerW}" height="${markerH + 10}">
            <rect x="1" y="1" width="${markerW - 2}" height="${markerH - 2}"
                  rx="4" ry="4"
                  fill="${fillColor}"
                  stroke="rgba(0,0,0,0.4)"
                  stroke-width="1.5"/>
            <text x="${markerW / 2}" y="${markerH / 2 + fontSize / 3}"
                  text-anchor="middle"
                  font-family="Arial,sans-serif"
                  font-size="${fontSize}px"
                  font-weight="bold"
                  fill="${fontColor}">${labelText}</text>
            <polygon points="${markerW/2 - 5},${markerH - 1} ${markerW/2 + 5},${markerH - 1} ${markerW/2},${markerH + 8}"
                     fill="${fillColor}"
                     stroke="rgba(0,0,0,0.4)"
                     stroke-width="1"/>
        </svg>`;

        const svgEncoded = "data:image/svg+xml;charset=UTF-8," + encodeURIComponent(svg);

        const marker = new google.maps.Marker({
            position: position,
            map:      map,
            icon: {
                url:         svgEncoded,
                scaledSize:  new google.maps.Size(markerW, markerH + 10),
                anchor:      new google.maps.Point(markerW / 2, markerH + 10)
            },
            title: row.address + " (" + row.label + ")"
        });

        const infoWindow = new google.maps.InfoWindow({
            content: "<strong>" + row.label + "</strong><br>" + row.address
        });
        marker.addListener("click", function() {
            infoWindow.open(map, marker);
        });

        markers.push(marker);
        bounds.extend(position);
        hasVisibleMarkers = true;
    });

    // Set initial view
    if (lastZoom && lastCenterLat && lastCenterLng) {
        map.setZoom(lastZoom);
        map.setCenter({ lat: lastCenterLat, lng: lastCenterLng });
    } else if (centerSet && centerLatLng) {
        map.setCenter(centerLatLng);
    } else if (hasVisibleMarkers) {
        map.fitBounds(bounds);
    }

    document.getElementById("btnCapture").disabled = false;
    document.getElementById("btnLoadMap").disabled = false;
    document.getElementById("btnReset").disabled   = false;
    setStatus("Ready — zoom or pan then click Capture");
}

function resetView() {
    if (!map || !mapData) return;

    lastZoom      = null;
    lastCenterLat = null;
    lastCenterLng = null;

    const bounds          = new google.maps.LatLngBounds();
    let hasVisibleMarkers = false;
    let centerSet         = false;
    let centerLatLng      = null;

    mapData.forEach(function(row) {
        const colDVal  = (row.flag || "").toLowerCase().trim();
        const hideAll  = colDVal.indexOf("hide all") > -1;
        const hidePin  = hideAll ||
                         colDVal === "yes" ||
                         colDVal === "y"   ||
                         colDVal.indexOf("hide") > -1;
        const isCenter = colDVal.indexOf("center") > -1;

        if (isCenter && !centerSet) {
            centerLatLng = { lat: row.lat, lng: row.lng };
            centerSet    = true;
        }
        if (hidePin) return;
        bounds.extend(new google.maps.LatLng(row.lat, row.lng));
        hasVisibleMarkers = true;
    });

    if (centerSet && centerLatLng) {
        map.setCenter(centerLatLng);
    } else if (hasVisibleMarkers) {
        map.fitBounds(bounds);
    }

    setStatus("View reset — zoom or pan then click Capture");
}

function captureMap() {
    setStatus("Capturing map...");
    document.getElementById("btnCapture").disabled = true;

    // Save current zoom and center
    const currentZoom   = map.getZoom();
    const currentCenter = map.getCenter();
    lastZoom      = currentZoom;
    lastCenterLat = currentCenter.lat();
    lastCenterLng = currentCenter.lng();

    const mapDiv = document.getElementById("map");

    if (typeof html2canvas === "undefined") {
        const script   = document.createElement("script");
        script.src     = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
        script.onload  = function() { doCapture(mapDiv); };
        script.onerror = function() {
            setStatus("Error: Could not load capture library.");
            document.getElementById("btnCapture").disabled = false;
        };
        document.head.appendChild(script);
    } else {
        doCapture(mapDiv);
    }
}

function doCapture(mapDiv) {
    html2canvas(mapDiv, {
        useCORS:    true,
        allowTaint: false,
        scale:      2
    }).then(function(canvas) {
        const base64 = canvas.toDataURL("image/png").split(",")[1];
        sendToExcel(base64);
    }).catch(function(err) {
        setStatus("Capture failed: " + err.message);
        document.getElementById("btnCapture").disabled = false;
    });
}

function sendToExcel(base64) {
    setStatus("Sending to Excel...");

    // Split into 30000 char chunks
    const chunkSize = 30000;
    const chunks    = [];
    for (let i = 0; i < base64.length; i += chunkSize) {
        chunks.push(base64.substring(i, i + chunkSize));
    }

    Excel.run(function(context) {
        var sheet       = context.workbook.sheets.getItem("Temp");
        var readyRange  = context.workbook.names.getItem("GMAP_READY").getRange();
        var chunksRange = context.workbook.names.getItem("GMAP_CHUNKS").getRange();
        var zoomRange   = context.workbook.names.getItem("GMAP_LAST_ZOOM").getRange();
        var latRange    = context.workbook.names.getItem("GMAP_LAST_LAT").getRange();
        var lngRange    = context.workbook.names.getItem("GMAP_LAST_LNG").getRange();

        // Write zoom and center
        zoomRange.values   = [[String(lastZoom)]];
        latRange.values    = [[String(lastCenterLat)]];
        lngRange.values    = [[String(lastCenterLng)]];

        // Write chunk count and ready flag
        chunksRange.values = [[chunks.length]];
        readyRange.values  = [["1"]];

        // Write chunks to Temp sheet starting at ROW_IMG_START (row 6, 0-based = 5)
        for (let i = 0; i < chunks.length; i++) {
            sheet.getCell(5 + i, 1).values = [[chunks[i]]];
        }

        return context.sync().then(function() {
            setStatus("Done — click Import Map Image in Excel ribbon to embed");
            document.getElementById("btnCapture").disabled = false;
        });

    }).catch(function(err) {
        setStatus("Error sending to Excel: " + err.message);
        document.getElementById("btnCapture").disabled = false;
    });
}