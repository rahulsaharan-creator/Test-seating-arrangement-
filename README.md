# Test Seating Arrangement

A website that creates the test seating arrangement Excel file for Matrix students. It produces the same sheets as the 4-Nov example.

Everything runs in the browser. The admission list is never uploaded to any server.

## How the operator uses it

The start screen has one button, **Create Seating Plan**. Then:

1. **Admission list:** upload the latest admission list (.xlsx). Only students marked Active are used. Optional: the room list, monitoring and due fee students, seating rules and the phase list.
2. **Test & papers:** the date, PRT or RPT, the courses, then each paper: name, 1 or 2 shifts, phases and batches.
3. **Online / Offline:** Online, Offline or Both for each paper (with Both, each batch or single students), and the top batches. **Create seating** picks the labs and buildings.
4. **Online seating** and **Offline seating:** check or change the labs, buildings, rooms and capacity for this test.
5. **Seating:** review the rooms (cards or table), reshuffle if wanted, and download the Excel file or the notice board.

A summary panel on the right (a bottom sheet on phones) updates as you choose. Back never loses data.

## Seating rules

- Students start in the building where their batch has class (from the batch–building list).
- Students from all chosen batches are shuffled together. Rooms are filled one at a time, and a room must be full before the next room is used.
- If the class building is full, the extra students move to the next near building in the list. Far buildings (PC by default) are used only after every near building is full. PC students who don't fit in PC move to the near buildings.
- Online batches use computer labs: the lab in their own building first, then the others.
- Inside each room, seat numbers follow roll-number order.
- The phase is the first letter of the batch (XB1 → X-Phase).

## Output sheets

Raw_Data, Student_list, Seating Arrangement, Attendance_1Shift (and Attendance_2Shift), SMD, Room_Layout, Master_List, and Not_Seated (only if someone could not get a seat).

## Built-in data

`src/data/defaults.json` holds the room list, lab list and batch–building list. The operator can also upload new files from the first screen (under "Update room list or batch–building list"). Uploads are saved in that browser.

## Development

```bash
npm install
npm run dev     # local site at http://localhost:5173
npm run build   # production build into dist/
npm test        # end-to-end check against the real Excel files (set DATA=/path/to/folder)
```

## Deploy on Vercel

Import this GitHub repo in Vercel. Vercel detects the Vite framework, so no settings need changing. Every push to `main` deploys automatically.
