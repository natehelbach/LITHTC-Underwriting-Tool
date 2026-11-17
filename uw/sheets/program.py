"""Program: physical build program (floor areas by use, parking, units by type)."""
from uw import xl
from uw.engine import Sheet, register


@register
class Program(Sheet):
    name = "Program"

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        site_sf = xl.n(ctx.ref("Initial Screen", "G18"))
        o["C8"] = site_sf / 43560
        o["C9"] = o["C8"] * 43560

        def override(key, default):
            v = xl.n(mi[key])
            return v if v > 0 else default

        lot = xl.n(ctx.ref("Initial Screen", "P17"))
        pk_floors = xl.n(ctx.ref("Initial Screen", "P36"))
        res_sf_screen = xl.n(ctx.ref("Initial Screen", "P40"))

        o["C13"] = override("rooftop_sf", 3000)
        o["C14"] = override("amenity_sf", 3000)
        o["C15"] = override("residential_sf_override",
                            res_sf_screen - o["C14"] + xl.n(mi["residential_sf_adjustment"]))

        # Share of each podium floor that is non-parking (from the overrides, else 0).
        def share(nonpk, pk):
            tot = xl.n(mi[pk]) + xl.n(mi[nonpk])
            return xl.n(mi[nonpk]) / tot if tot else 0.0

        o["D17"] = share("4th_floor_non_parking_sf", "4th_floor_parking_sf")
        o["D19"] = share("3rd_floor_non_parking_sf", "3rd_floor_parking_sf")
        o["D21"] = share("2nd_floor_non_parking_sf", "2nd_floor_parking_sf")
        tot_g = xl.n(mi["ground_floor_parking_sf"]) + xl.n(mi["ground_floor_non_parking_sf"])
        o["D23"] = xl.n(mi["ground_floor_non_parking_sf"]) / tot_g if tot_g else 0.225

        # A partial top parking level (e.g. 2.5 floors) is spread over the floors.
        frac = xl.ROUNDUP(pk_floors, 0) - pk_floors
        spread = 2 if 2 < pk_floors < 3 else 3 if 3 < pk_floors < 4 else 4 if 4 < pk_floors < 5 else 1

        def nonparking(level_above, share_):
            return lot * share_ * (1 - frac) if pk_floors > level_above else 0

        def parking(level_above, nonpk_sf):
            return lot * (1 - frac / spread) - nonpk_sf if pk_floors > level_above else 0

        o["C17"] = override("4th_floor_non_parking_sf", nonparking(3, o["D17"]))
        o["C16"] = override("4th_floor_parking_sf", parking(3, o["C17"]))
        o["C19"] = override("3rd_floor_non_parking_sf", nonparking(2, o["D19"]))
        o["C18"] = override("3rd_floor_parking_sf", parking(2, o["C19"]))
        o["C21"] = override("2nd_floor_non_parking_sf", nonparking(1, o["D21"]))
        o["C20"] = override("2nd_floor_parking_sf", parking(1, o["C21"]))
        o["C23"] = override("ground_floor_non_parking_sf", lot * o["D23"])
        o["C22"] = override("ground_floor_parking_sf", lot - o["C23"])
        o["C24"] = sum(o[f"C{r}"] for r in range(14, 24))

        # Building area. NOTE (from Excel): common/parking area counts only the ground
        # and 2nd floors (C20:C22); 3rd/4th floor parking is excluded. Kept for parity.
        o["C29"] = o["C15"] + o["C14"] + o["C23"]
        o["C30"] = o["C22"] + o["C21"] + o["C20"]
        o["C31"] = o["C30"] + o["C29"]
        o["C28"] = o["C30"] / o["C31"]

        # Units by type (from Rent Builder).
        o.update({"B52": "Manager", "D52": "AVG SF", "E52": "# of units"})  # header inside E46:E56
        types = ["Studio", "1 BR", "2 BR", "3 BR"]
        for i, t in enumerate(types):
            o[f"B{46 + i}"] = t
            o[f"D{46 + i}"] = ctx.ref("Rent Builder", f"D{53 + i}")
            o[f"E{46 + i}"] = ctx.ref("Rent Builder", f"E{53 + i}")
            o[f"D{53 + i}"] = ctx.ref("Rent Builder", f"D{60 + i}")
            o[f"E{53 + i}"] = ctx.ref("Rent Builder", f"E{60 + i}")
            o[f"D{60 + i}"] = ctx.ref("Rent Builder", f"D{68 + i}")
            o[f"E{60 + i}"] = ctx.ref("Rent Builder", f"E{68 + i}")
        rb_type = ctx.col("Rent Builder", "H12:H39")
        rb_cnt = ctx.col("Rent Builder", "F12:F39")
        o["H47"] = xl.SUMIF(rb_type, "1 BR", rb_cnt)
        o["H48"] = xl.SUMIF(rb_type, "2 BR", rb_cnt)
        for base, last in ((46, 50), (53, 57), (60, 64)):
            rows = range(base, base + 4)
            o[f"D{last}"] = xl.SUMPRODUCT([o[f"D{r}"] for r in rows], [o[f"E{r}"] for r in rows])
            o[f"E{last}"] = xl.SUM([o[f"E{r}"] for r in rows])

        # Parking.
        stalls = xl.n(ctx.ref("Unit Mix Summary", "F36"))
        o["C36"] = stalls
        o["C38"] = o["C20"] + o["C22"]
        o["C35"] = xl.div(o["C38"], stalls, 0.0)
        o["C39"] = xl.div(stalls, o["E64"], 0.0)
        o["C41"] = o["C38"] / o["C29"]
        return o
