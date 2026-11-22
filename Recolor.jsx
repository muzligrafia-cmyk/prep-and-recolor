// Adobe Illustrator Script: Pink-Green Recolour (V2.1)
// Unconditional recolour. No colour detection. No menu commands.
//
//   Compound paths -> sub-paths extracted directly via DOM (no menu commands,
//                     no selection churn), wrapped in a sub-group placed at
//                     the compound's original z-position.
//   Then every path - loose or from a compound - gets a nesting depth =
//   how many closed shapes it sits inside. Fill removed.
//     PINK  = even depth (0, 2, 4...) AND has something inside it
//     GREEN = everything else (holes, lone shapes, open lines)
//   So: part = pink, its holes = green, a part inside a hole = pink if it
//   has holes of its own (else green), its holes = green, and so on.
//   Same rule as Prep.js V2.1 and Audit.js V3.5.
//
// Existing layers and groups are preserved.
// Single-path compounds are unwrapped in place (no needless group).
// If extraction fails, the compound is left intact and stroked pink.
// Items on the "duplicates" and "skipped" layers are ignored.
//
// V2.0: Inside/outside detection (same rules as Prep.js) instead of
//       "largest piece = outer" and "all simple paths green".
// V2.1: Depth rule shared with Prep.js and Audit.js: compound pieces and
//       loose paths are judged together, nested parts alternate pink/green,
//       outer pieces with nothing inside are green.

function pinkGreenRecolour() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    if (doc.selection.length === 0) {
        alert("Select the artwork to process.");
        return;
    }

    var IGNORE_LAYERS = ["duplicates", "skipped"];
    var PT_PER_MM  = 2.834645669;
    var EDGE_TOL   = 0.05 * PT_PER_MM;  // points this close to an outline count as "on" it
    var CURVE_STEP = 0.25 * PT_PER_MM;  // how finely curves are measured

    var items = [];
    for (var s = 0; s < doc.selection.length; s++) items.push(doc.selection[s]);

    var cPink  = mkRGB(255, 0, 255);
    var cGreen = mkRGB(0, 255, 0);

    function mkRGB(r, g, b) {
        var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c;
    }

    var allInfos   = []; // geometry of every path touched (for inside/outside tests)
    var candidates = []; // loose simple paths, coloured at the end: {item, info}

    // Suppress dialogs during the run; always restored in finally.
    var prevInteraction = app.userInteractionLevel;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    try {
        for (var i = items.length - 1; i >= 0; i--) {
            try { processItem(items[i]); } catch (e) {}
        }
        colourCandidates();
        doc.selection = null;
    } finally {
        app.userInteractionLevel = prevInteraction;
    }

    // =========================================================================

    function processItem(item) {
        if (isIgnored(item)) return;

        if (item.typename === "TextFrame") {
            try { item = item.createOutline(); } catch (e) { return; }
        }
        if (item.typename === "GroupItem") {
            // Snapshot children first: splitting inserts new sub-groups into
            // this parent, which would shift a live index-based loop.
            var kids = [];
            for (var g = 0; g < item.pageItems.length; g++) kids.push(item.pageItems[g]);
            for (var k = kids.length - 1; k >= 0; k--) {
                try { processItem(kids[k]); } catch (e) {}
            }
            return;
        }
        if (item.typename === "PathItem") {
            try { if (item.guides) return; } catch (e) {}
            // Colour decided at the end, once every shape is known
            candidates.push({ item: item, info: addInfo(item) });
            return;
        }
        if (item.typename === "CompoundPathItem") {
            splitAndGroup(item);
            return;
        }
    }

    function isIgnored(item) {
        try {
            var name = item.layer.name;
            for (var i = 0; i < IGNORE_LAYERS.length; i++) {
                if (name === IGNORE_LAYERS[i]) return true;
            }
        } catch (e) {}
        return false;
    }

    // --- COMPOUND PATHS ------------------------------------------------------

    function splitAndGroup(compound) {
        var total = 0;
        try { total = compound.pathItems.length; } catch (e) {}

        if (total === 0) {
            // Empty shell: nothing to recolour, just remove it.
            try { compound.remove(); } catch (e) {}
            return;
        }

        // Single sub-path: unwrap in place, no group needed.
        // Colour decided at the end like any loose path.
        if (total === 1) {
            try {
                var only = compound.pathItems[0];
                only.move(compound, ElementPlacement.PLACEBEFORE);
                compound.remove();
                candidates.push({ item: only, info: buildInfo(only) });
            } catch (e) {
                try { applyTo(compound, cPink); } catch (e2) {}
            }
            return;
        }

        // Multi-path compound: create the target group at the compound's
        // original z-position, then move sub-paths out via the DOM.
        var group = null;
        try {
            group = compound.parent.groupItems.add();
            group.move(compound, ElementPlacement.PLACEBEFORE);
        } catch (e) {
            try { applyTo(compound, cPink); } catch (e2) {}
            return;
        }

        var released = [];
        // Iterate backwards: moving an item shrinks the live collection.
        for (var p = total - 1; p >= 0; p--) {
            try {
                var path = compound.pathItems[p];
                path.move(group, ElementPlacement.PLACEATBEGINNING);
                released.push(path);
            } catch (e) {}
        }

        if (released.length === 0) {
            try { group.remove(); } catch (e) {}
            try { applyTo(compound, cPink); } catch (e2) {}
            return;
        }

        try {
            if (compound.pathItems.length === 0) compound.remove();
            else applyTo(compound, cPink); // partial extraction: keep remainder visible
        } catch (e) {}

        // Pink or green is decided at the end, together with every other
        // path in the selection (see colourCandidates)
        for (var m = 0; m < released.length; m++) {
            candidates.push({ item: released[m], info: buildInfo(released[m]) });
        }
    }

    // --- LOOSE SIMPLE PATHS: PINK OR GREEN -----------------------------------

    function colourCandidates() {
        for (var c = 0; c < candidates.length; c++) {
            var cand = candidates[c];
            var pink = false;

            // Depth rule (same as Audit.js): pink = even depth AND has something inside
            if (cand.info) {
                var depth = 0, hasInside = false;
                for (var j = 0; j < candidates.length; j++) {
                    if (j === c) continue;
                    var other = candidates[j].info;
                    if (!other) continue;
                    if (isInside(cand.info, other)) depth++;
                    else if (!hasInside && isInside(other, cand.info)) hasInside = true;
                }
                pink = (depth % 2 === 0) && hasInside;
            }

            try { applyTo(cand.item, pink ? cPink : cGreen); } catch (e) {}
        }
    }

    function applyTo(p, color) {
        try { p.filled = false; } catch (e) {}
        try { p.stroked = true; } catch (e) {}
        try { p.strokeColor = color; } catch (e) {}
        try { p.strokeWidth = 1; } catch (e) {}
    }

    // --- GEOMETRY: what's inside what ----------------------------------------

    function addInfo(path) {
        var inf = buildInfo(path);
        if (inf) allInfos.push(inf);
        return inf;
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

pinkGreenRecolour();
