/* global Office, Excel, google, html2canvas */

let map           = null;
let markers       = [];
let apiKey        = null;
let mapId         = null;
let mapData       = null;
let lastZoom      = null;
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
    document.getElementById("btnLoadMap").disabled = true;
    document.getElementById("btnCapture").disabled = true;
    document.getElementById("btnReset").disabled   = true;

    Excel.run(function(context) {
        var sheet      = context.workbook.worksheets.getItem("Temp");
        var apiKeyCell = sheet.getCell(0, 1);
        var mapIdCell  = sheet.getCell(1, 1);
        var dataCell   = sheet.getCell(2, 1);
        var zoomCell   = sheet.getCell(0, 2);
        var latCell    = sheet.getCell(0, 3);
        var lngCell    = sheet.getCell(0, 4);
        var widthCell  = sheet.getCell(0, 5);  // F1
        var heightCell = sheet.getCell(0, 6);  // G1
        

        apiKeyCell.load("values");
        mapIdCell.load("values");
        dataCell.load("values");
        zoomCell.load("values");
        latCell.load("values");
        lngCell.load("values");
        widthCell.load("values");
        heightCell.load("values");        
        
        return context.sync().then(function() {
            apiKey = apiKeyCell.values[0][0];
            mapId  = mapIdCell.values[0][0] || "";
            var raw       = dataCell.values[0][0];
            var savedZoom = zoomCell.values[0][0];
            var savedLat  = latCell.values[0][0];
            var savedLng  = lngCell.values[0][0];
            var imgWidth  = widthCell.values[0][0]  || 400;
            var imgHeight = heightCell.values[0][0] || 400;
            window._gmapWidth  = imgWidth;
            window._gmapHeight = imgHeight;

            if (savedZoom && savedZoom !== "") {
                lastZoom      = parseInt(savedZoom);
                lastCenterLat = parseFloat(savedLat);
                lastCenterLng = parseFloat(savedLng);
            }

            if (!apiKey || !raw) {
                setStatus("Error: No data. Run Open_GoogleMap_Addin_TaskPane macro first.");
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
    // Only load Google Maps once
    if (typeof google !== "undefined" && typeof google.maps !== "undefined") {
        window._gmapMapId = mid;
        initMap();
        return;
    }

    const existing = document.getElementById("gmaps-script");
    if (existing) existing.remove();

    setStatus("Loading Google Maps...");
    window._gmapMapId = mid;

    const script   = document.createElement("script");
    script.id      = "gmaps-script";
    script.src     = "https://maps.googleapis.com/maps/api/js?key=" + key + "&callback=initMap";
    script.async   = true;
    script.defer   = true;
    script.onerror = function() {
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
        zoom:             lastZoom || 12,
        center:           { lat: lastCenterLat || 0, lng: lastCenterLng || 0 },
        mapTypeId:        "roadmap",
        fullscreenControl: false,
        streetViewControl: false,
        mapTypeControl:   true,
        zoomControl:      true,
        renderingType:    google.maps.RenderingType.RASTER
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

        if (isCenter && !centerSet) {
            centerLatLng = { lat: row.lat, lng: row.lng };
            centerSet    = true;
        }

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
                url:        svgEncoded,
                scaledSize: new google.maps.Size(markerW, markerH + 10),
                anchor:     new google.maps.Point(markerW / 2, markerH + 10)
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

    lastZoom      = map.getZoom();
    lastCenterLat = map.getCenter().lat();
    lastCenterLng = map.getCenter().lng();

    // Resize map div to exact img_rng dimensions
    // Excel points to pixels: 1 point = 1.333 pixels at 96dpi
    const pxWidth  = Math.round((window._gmapWidth  || 400) * 1.333);
    const pxHeight = Math.round((window._gmapHeight || 400) * 1.333);

    const mapDiv = document.getElementById("map");
    mapDiv.style.width  = pxWidth  + "px";
    mapDiv.style.height = pxHeight + "px";

    // Trigger map resize event
    google.maps.event.trigger(map, "resize");

    // Restore center after resize
    map.setCenter({ lat: lastCenterLat, lng: lastCenterLng });
    map.setZoom(lastZoom);

    // Wait for map to re-render at new size
    setTimeout(doCapture, 1000);
}

function doCapture() {
    const mapDiv = document.getElementById("map");
    setStatus("Processing image...");

    html2canvas(mapDiv, {
        useCORS:         true,
        allowTaint:      true,
        scale:           1,
        logging:         false,
        imageTimeout:    0,
        removeContainer: false,
        width:           mapDiv.offsetWidth,
        height:          mapDiv.offsetHeight
    }).then(function(canvas) {
        const base64 = canvas.toDataURL("image/png").split(",")[1];
        console.log("Capture success - base64 length:", base64.length);

        // Restore map div to full size
        mapDiv.style.width  = "100%";
        mapDiv.style.height = "100%";
        google.maps.event.trigger(map, "resize");

        writeChunksToExcel(base64);
    }).catch(function(err) {
        console.log("Capture error:", err);
        // Restore map div on error too
        mapDiv.style.width  = "100%";
        mapDiv.style.height = "100%";
        setStatus("Capture failed: " + (err && err.message ? err.message : String(err)));
        document.getElementById("btnCapture").disabled = false;
    });
}
function writeChunksToExcel(base64) {
    setStatus("Sending to Excel...");

    // Split base64 into 30000 char chunks
    const chunkSize = 30000;
    const chunks    = [];
    for (let i = 0; i < base64.length; i += chunkSize) {
        chunks.push(base64.substring(i, i + chunkSize));
    }

    console.log("Total chunks:", chunks.length);

    Excel.run(function(context) {
        var sheet = context.workbook.worksheets.getItem("Temp");

        // C1, D1, E1 — zoom and center (0-based: row 0, cols 2,3,4)
        sheet.getCell(0, 2).values = [[String(lastZoom)]];
        sheet.getCell(0, 3).values = [[String(lastCenterLat)]];
        sheet.getCell(0, 4).values = [[String(lastCenterLng)]];

        // B4 — ready flag (0-based: row 3, col 1)
        sheet.getCell(3, 1).values = [["1"]];

        // B5 — chunk count (0-based: row 4, col 1)
        sheet.getCell(4, 1).values = [[chunks.length]];

        // B6 onwards — image chunks (0-based: row 5+, col 1)
        for (let i = 0; i < chunks.length; i++) {
            sheet.getCell(5 + i, 1).values = [[chunks[i]]];
        }

        return context.sync().then(function() {
            setStatus("Done — click Import Map Image in Excel ribbon to embed");
            document.getElementById("btnCapture").disabled = false;
        });

    }).catch(function(err) {
        setStatus("Error writing to Excel: " + err.message);
        document.getElementById("btnCapture").disabled = false;
    });
}