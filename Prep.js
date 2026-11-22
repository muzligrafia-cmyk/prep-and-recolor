// Adobe Illustrator Script: Smart Laser Recolour (V2.1)
// Standardises stroke colours for laser cutting. Items stay in their layers
// and groups - except skipped items, which go to the "skipped" layer.
//
// RULES (colours must be EXACT - the RIP software is binary)
//   Black fill (#000)  -> outline removed, shape otherwise untouched (engraving)
//   Black stroke       -> restyled black
//   Red stroke         -> restyled red
//   Blue stroke        -> restyled blue
//   Everything else is a CUT line. Compound cut paths are released first
//   (pieces wrapped in a new group inside the original parent so they still
//   move together). Then every cut line - loose or from a compound - gets a
//   nesting depth = how many closed cut shapes it sits inside:
//     PINK  = even depth (0, 2, 4...) AND has a cut line inside it
//     GREEN = everything else (holes, lone shapes, open lines)
//   So: part = pink, its holes = green, a part inside a hole = pink if it
//   has holes of its own (else green), its holes = green, and so on.
//   Only cut lines count for this - engraving, scoring and perforation
//   inside a part don't make it pink. Same rule as Audit.js V3.5.
//   (Pink/green paths from an earlier run stay correct if re-run.)
//
//   Items on the "duplicates" and "skipped" layers are ignored completely.
//   Images, symbols, blends/effect objects, meshes, graphs, and anything the
//   script fails to process -> moved to the "skipped" layer, with an alert
//   explaining why. No alert if nothing was skipped.
//
// V2.0: Inside/outside detection instead of "largest shape = outer",
//       containment for loose simple paths, strict #000 black, black fills
//       always lose their outline, "skipped" layer + report, ignores the
//       "duplicates" layer.
// V2.1: Depth rule shared with Audit.js and Recolor.jsx: compound pieces and
//       loose paths are judged together, nested parts alternate pink/green,
//       outer pieces with nothing inside are green, only cut lines count.

function smartLaserRecolour() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    if (doc.selection.length === 0) {
        alert("Select the artwork to process.");
        return;
    }

    // --- Settings ---
    var IGNORE_LAYERS   = ["duplicates", "skipped"];
    var SKIP_LAYER_NAME = "skipped";
    var PT_PER_MM  = 2.834645669;
    var EDGE_TOL   = 0.05 * PT_PER_MM;  // points this close to an outline count as "on" it
    var CURVE_STEP = 0.25 * PT_PER_MM;  // how finely curves are measured

    // Snapshot selection - processing modifies doc.selection
    var items = [];
    for (var s = 0; s < doc.selection.length; s++) items.push(doc.selection[s]);

    // Trigger colours
    var cPink  = mkRGB(255, 0,   255);
    var cGreen = mkRGB(0,   255, 0);
    var cRed   = mkRGB(255, 0,   0);
    var cBlue  = mkRGB(0,   0,   255);
    var cBlack = mkRGB(0,   0,   0);

    function mkRGB(r, g, b) {
        var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c;
    }

    var laserInfos    = []; // geometry of every path in the job (for inside/outside tests)
    var cutCandidates = []; // loose simple cut paths, coloured at the end: {item, info}
    var skipped       = []; // {item, reason}

    // --- Run ---
    for (var i = items.length - 1; i >= 0; i--) {
        try { processItem(items[i]); }
        catch (e) { skipItem(items[i], "Unexpected error while processing"); }
    }
    doc.selection = null;

    colourCutCandidates();
    moveSkippedAndReport();

    // =========================================================================

    function processItem(item) {
        if (isIgnored(item)) return;
        var t = item.typename;

        if (t === "TextFrame") {
            var outlined;
            try { outlined = item.createOutline(); }
            catch (e) { skipItem(item, "Text that couldn't be converted to outlines"); return; }
            processItem(outlined);
            return;
        }
        if (t === "GroupItem") {
            // Snapshot children first: splitting inserts new sub-groups into
            // this parent, which would shift a live index-based loop.
            var kids = [];
            for (var g = 0; g < item.pageItems.length; g++) kids.push(item.pageItems[g]);
            for (var k = kids.length - 1; k >= 0; k--) {
                try { processItem(kids[k]); }
                catch (e) { skipItem(kids[k], "Unexpected error while processing"); }
            }
            return;
        }
        if (t === "PathItem") {
            try { if (item.guides) return; } catch (e) { }
            routeSimplePath(item);
            return;
        }
        if (t === "CompoundPathItem") {
            routeCompound(item);
            return;
        }
        skipItem(item, describeType(t));
    }

    function isIgnored(item) {
        try {
            var name = item.layer.name;
            for (var i = 0; i < IGNORE_LAYERS.length; i++) {
                if (name === IGNORE_LAYERS[i]) return true;
            }
        } catch (e) { }
        return false;
    }

    function describeType(t) {
        if (t === "PlacedItem")     return "Linked image or file (not laser vector)";
        if (t === "RasterItem")     return "Embedded image (not laser vector)";
        if (t === "SymbolItem")     return "Symbol (expand it first)";
        if (t === "PluginItem")     return "Blend, envelope or other effect object (expand it first)";
        if (t === "MeshItem")       return "Gradient mesh";
        if (t === "GraphItem")      return "Graph";
        if (t === "LegacyTextItem") return "Legacy text";
        if (t === "NonNativeItem")  return "Non-native artwork";
        return "Unsupported object (" + t + ")";
    }

    function skipItem(item, reason) {
        skipped.push({ item: item, reason: reason });
    }

    // --- SIMPLE PATH ---------------------------------------------------------

    function routeSimplePath(item) {
        // Black fill = engraving: remove any outline, otherwise leave alone
        if (hasBlackFill(item)) {
            stripOutline(item);
            addInfo(item);
            return;
        }

        var stroke = safeStroke(item);
        if (stroke) {
            if (isBlack(stroke))              { applyTo(item, cBlack); addInfo(item); return; }
            if (isExactRGB(stroke, 255, 0, 0)) { applyTo(item, cRed);   addInfo(item); return; }
            if (isExactRGB(stroke, 0, 0, 255)) { applyTo(item, cBlue);  addInfo(item); return; }
        }

        // Cut line - pink or green is decided at the end, once every
        // shape is known (see colourCutCandidates)
        cutCandidates.push({ item: item, info: addInfo(item) });
    }

    // --- COMPOUND PATH -------------------------------------------------------

    function routeCompound(item) {
        var first = readFirstPath(item);

        // Shape Builder compound - normalize to make sub-paths readable
        if (!first) {
            item = normalizeCompound(item);
            if (!item) return;
            first = readFirstPath(item);
        }

        if (first) {
            if (hasBlackFill(first)) {
                stripOutline(item);
                addSubpathInfos(item);
                return;
            }
            var stroke = safeStroke(first);
            if (stroke) {
                if (isBlack(stroke))              { styleItem(item, cBlack); addSubpathInfos(item); return; }
                if (isExactRGB(stroke, 255, 0, 0)) { styleItem(item, cRed);   addSubpathInfos(item); return; }
                if (isExactRGB(stroke, 0, 0, 255)) { styleItem(item, cBlue);  addSubpathInfos(item); return; }
            }
        }

        // Anything else -> release, then outer pieces pink, inner pieces green
        splitAndGroup(item);
    }

    function readFirstPath(item) {
        try { if (item.pathItems.length > 0) return item.pathItems[0]; } catch (e) { }
        return null;
    }

    function safeStroke(p) {
        try {
            if (!p.stroked) return null;
            var c = p.strokeColor;
            if (!c || c.typename === "NoColor") return null;
            return c;
        } catch (e) { return null; }
    }

    function hasBlackFill(p) {
        try { return p.filled && isBlack(p.fillColor); } catch (e) { return false; }
    }

    // --- NORMALIZE (release -> re-compound) ----------------------------------
    // Only used to make a Shape Builder compound's sub-paths readable.

    function normalizeCompound(item) {
        doc.selection = null;
        item.selected = true;

        try { app.executeMenuCommand("noCompoundPath"); }
        catch (e) { doc.selection = null; return item; }

        try { app.executeMenuCommand("compoundPath"); }
        catch (e) {
            // Re-compound failed - treat the released paths as loose paths
            var rel = [];
            for (var i = 0; i < doc.selection.length; i++) rel.push(doc.selection[i]);
            doc.selection = null;
            for (var r = 0; r < rel.length; r++) {
                try {
                    if (rel[r].typename === "PathItem") routeSimplePath(rel[r]);
                } catch (e2) { skipItem(rel[r], "Unexpected error while processing"); }
            }
            return null;
        }

        var result = (doc.selection.length > 0) ? doc.selection[0] : null;
        doc.selection = null;
        return result;
    }

    // --- SPLIT & GROUP -------------------------------------------------------

    function splitAndGroup(item) {
        var parent = null;
        try { parent = item.parent; } catch (e) { }

        doc.selection = null;
        item.selected = true;

        try { app.executeMenuCommand("noCompoundPath"); }
        catch (e) {
            doc.selection = null;
            skipItem(item, "Compound path that couldn't be released");
            return;
        }

        var released = [];
        for (var i = 0; i < doc.selection.length; i++) {
            if (doc.selection[i].typename === "PathItem") released.push(doc.selection[i]);
        }
        doc.selection = null;

        if (released.length === 0) {
            skipItem(item, "Compound path that couldn't be released");
            return;
        }

        // Pink or green is decided at the end, together with every other
        // cut line in the selection (see colourCutCandidates)
        for (var m = 0; m < released.length; m++) {
            cutCandidates.push({ item: released[m], info: buildInfo(released[m]) });
        }

        // Wrap multiple pieces in a new group inside the original parent
        if (released.length > 1 && parent) {
            var group = null;
            try { group = parent.groupItems.add(); } catch (e) { }
            if (group) {
                for (var k = 0; k < released.length; k++) {
                    try { released[k].move(group, ElementPlacement.PLACEATEND); } catch (e) { }
                }
            }
        }
    }

    // --- LOOSE CUT PATHS: PINK OR GREEN --------------------------------------

    // Depth rule (same as Audit.js): pink = even depth AND has a cut line inside
    function colourCutCandidates() {
        for (var c = 0; c < cutCandidates.length; c++) {
            var cand = cutCandidates[c];
            var pink = false;

            if (cand.info) {
                var depth = 0, hasInside = false;
                for (var j = 0; j < cutCandidates.length; j++) {
                    if (j === c) continue;
                    var other = cutCandidates[j].info;
                    if (!other) continue;
                    if (isInside(cand.info, other)) depth++;
                    else if (!hasInside && isInside(other, cand.info)) hasInside = true;
                }
                pink = (depth % 2 === 0) && hasInside;
            }

            try { applyTo(cand.item, pink ? cPink : cGreen); } catch (e) { }
        }
    }

    // --- SKIPPED LAYER + REPORT ----------------------------------------------

    function moveSkippedAndReport() {
        if (skipped.length === 0) return; // Nothing skipped - no alert

        var layer;
        try {
            layer = doc.layers.getByName(SKIP_LAYER_NAME);
        } catch (e) {
            layer = doc.layers.add();
            layer.name = SKIP_LAYER_NAME;
        }
        layer.locked = false;
        layer.visible = true;

        var counts = {}, order = [], notMoved = 0;
        for (var i = 0; i < skipped.length; i++) {
            var sk = skipped[i];
            try { sk.item.move(layer, ElementPlacement.PLACEATBEGINNING); }
            catch (e) { notMoved++; }
            if (!counts[sk.reason]) { counts[sk.reason] = 0; order.push(sk.reason); }
            counts[sk.reason]++;
        }

        var msg = "Some items were skipped and moved to the '" + SKIP_LAYER_NAME + "' layer:\n";
        for (var r = 0; r < order.length; r++) {
            msg += "\n- " + counts[order[r]] + " x " + order[r];
        }
        if (notMoved > 0) {
            msg += "\n\n" + notMoved + " of them couldn't be moved (probably locked) and are still in place.";
        }
        alert(msg);
    }

    // --- STYLE ---------------------------------------------------------------

    function styleItem(item, color) {
        applyTo(item, color);
        if (item.typename === "CompoundPathItem") {
            try {
                for (var i = 0; i < item.pathItems.length; i++) {
                    try { applyTo(item.pathItems[i], color); } catch (e) { }
                }
            } catch (e) { }
        }
    }

    function applyTo(p, color) {
        try { p.filled = false; } catch (e) { }
        try { p.stroked = true; } catch (e) { }
        try { p.strokeColor = color; } catch (e) { }
        try { p.strokeWidth = 1; } catch (e) { }
    }

    // Removes the outline from a black-filled engraving shape, keeps the fill
    function stripOutline(item) {
        if (item.typename === "CompoundPathItem") {
            try {
                for (var i = 0; i < item.pathItems.length; i++) {
                    try { item.pathItems[i].stroked = false; } catch (e) { }
                }
            } catch (e) { }
        } else {
            try { item.stroked = false; } catch (e) { }
        }
    }

    // --- COLOUR DETECTION (strict) -------------------------------------------
    // Values are rounded to whole numbers, i.e. exactly what the hex code shows.

    function isBlack(c) {
        if (!c) return false;
        if (c.typename === "SpotColor") {
            try { return isBlack(c.spot.color); } catch (e) { return false; }
        }
        if (c.typename === "RGBColor") {
            return Math.round(c.red) === 0 && Math.round(c.green) === 0 && Math.round(c.blue) === 0;
        }
        if (c.typename === "GrayColor") return Math.round(c.gray) === 100; // 100% gray = #000
        return false;
    }

    function isExactRGB(c, r, g, b) {
        if (!c) return false;
        if (c.typename === "SpotColor") {
            try { return isExactRGB(c.spot.color, r, g, b); } catch (e) { return false; }
        }
        return c.typename === "RGBColor" &&
               Math.round(c.red) === r && Math.round(c.green) === g && Math.round(c.blue) === b;
    }

    // --- GEOMETRY: what's inside what ----------------------------------------

    function addInfo(path) {
        var inf = buildInfo(path);
        if (inf) laserInfos.push(inf);
        return inf;
    }

    function addSubpathInfos(compound) {
        try {
            for (var i = 0; i < compound.pathItems.length; i++) addInfo(compound.pathItems[i]);
        } catch (e) { }
    }

    function dist(a, b) {
        var dx = a[0] - b[0], dy = a[1] - b[1];
        return Math.sqrt(dx * dx + dy * dy);
    }

    // Turns one curve segment into a chain of short straight pieces
    function flattenSegment(p0, p1, p2, p3, out) {
        if (dist(p0, p1) < 0.001 && dist(p2, p3) < 0.001) { out.push(p3[0], p3[1]); return; }
        var est = dist(p0, p1) + dist(p1, p2) + dist(p2, p3);
        var steps = Math.ceil(est / CURVE_STEP);
        if (steps < 4) steps = 4;
        if (steps > 200) steps = 200;
        for (var i = 1; i <= steps; i++) {
            var t = i / steps, mt = 1 - t;
            var a = mt * mt * mt, b = 3 * mt * mt * t, c = 3 * mt * t * t, d = t * t * t;
            out.push(a * p0[0] + b * p1[0] + c * p2[0] + d * p3[0],
                     a * p0[1] + b * p1[1] + c * p2[1] + d * p3[1]);
        }
    }

    function buildInfo(path) {
        var pts, n, closed;
        try { pts = path.pathPoints; n = pts.length; closed = path.closed; } catch (e) { return null; }
        if (n < 2) return null;

        var poly = [];
        var a0 = pts[0].anchor;
        poly.push(a0[0], a0[1]);
        var segCount = closed ? n : n - 1;
        for (var s = 0; s < segCount; s++) {
            var P = pts[s], Q = pts[(s + 1) % n];
            flattenSegment(P.anchor, P.rightDirection, Q.leftDirection, Q.anchor, poly);
        }

        var minX = poly[0], maxX = poly[0], minY = poly[1], maxY = poly[1];
        for (var i = 2; i < poly.length; i += 2) {
            var x = poly[i], y = poly[i + 1];
            if (x < minX) minX = x; if (x > maxX) maxX = x;
            if (y < minY) minY = y; if (y > maxY) maxY = y;
        }

        // Enclosed area (shoelace formula); open lines enclose nothing
        var area = 0;
        if (closed) {
            var m = poly.length / 2;
            for (var j = 0; j < m; j++) {
                var k = (j + 1) % m;
                area += poly[2 * j] * poly[2 * k + 1] - poly[2 * k] * poly[2 * j + 1];
            }
            area = Math.abs(area / 2);
        }

        return { poly: poly, closed: closed, area: area,
                 minX: minX, maxX: maxX, minY: minY, maxY: maxY };
    }

    function pointInPoly(x, y, poly) {
        var inside = false, n = poly.length / 2;
        for (var i = 0, j = n - 1; i < n; j = i++) {
            var xi = poly[2 * i], yi = poly[2 * i + 1], xj = poly[2 * j], yj = poly[2 * j + 1];
            if (((yi > y) !== (yj > y)) && (x < (xj - xi) * (y - yi) / (yj - yi) + xi)) inside = !inside;
        }
        return inside;
    }

    function nearEdge(x, y, poly) {
        var n = poly.length / 2, tolSq = EDGE_TOL * EDGE_TOL;
        for (var i = 0; i < n - 1; i++) {
            var x1 = poly[2 * i], y1 = poly[2 * i + 1], x2 = poly[2 * i + 2], y2 = poly[2 * i + 3];
            var dx = x2 - x1, dy = y2 - y1, len2 = dx * dx + dy * dy, t = 0;
            if (len2 > 0) {
                t = ((x - x1) * dx + (y - y1) * dy) / len2;
                if (t < 0) t = 0; else if (t > 1) t = 1;
            }
            var cx = x1 + t * dx - x, cy = y1 + t * dy - y;
            if (cx * cx + cy * cy <= tolSq) return true;
        }
        return false;
    }

    // Is shape a completely inside closed shape b?
    function isInside(a, b) {
        if (!a || !b || a === b || !b.closed) return false;
        if (a.minX < b.minX - EDGE_TOL || a.maxX > b.maxX + EDGE_TOL ||
            a.minY < b.minY - EDGE_TOL || a.maxY > b.maxY + EDGE_TOL) return false;
        if (b.area <= a.area) return false; // The container must be the bigger shape

        var p = a.poly, n = p.length / 2;
        var stride = Math.max(1, Math.floor(n / 200)); // Cap the work on very detailed shapes
        for (var i = 0; i < n; i += stride) {
            var x = p[2 * i], y = p[2 * i + 1];
            if (!pointInPoly(x, y, b.poly) && !nearEdge(x, y, b.poly)) return false;
        }
        return true;
    }
}

// Run via a deferred BridgeTalk message rather than a direct call.
// When this script is triggered from an Action (e.g. bound to a hotkey like F6),
// Illustrator's Action Player still holds its internal menu-command dispatch lock
// while the "Run Script" step executes. Any app.executeMenuCommand() calls made
// synchronously inside that step (e.g. "noCompoundPath" / "compoundPath", used
// above to split compound paths) can silently no-op or behave inconsistently.
//
// Sending the work as a BridgeTalk message to Illustrator queues it to run on the
// next idle cycle - after the Action Player finishes and releases that lock - so
// executeMenuCommand behaves the same as when run directly from File > Scripts.
// This is used instead of app.scheduleTask() because scheduleTask isn't available
// in every Illustrator build; BridgeTalk is supported much more broadly.
var bt = new BridgeTalk();
bt.target = "illustrator";
bt.body = "(" + smartLaserRecolour.toString() + ")();";
bt.onError = function (errObj) {
    alert("Smart Laser Recolour failed to run: " + errObj.body);
};
bt.send();
