# Test Seating Arrangement

A website that creates the test seating arrangement Excel file for Matrix students. It produces the same sheets as the 4-Nov example.

Everything runs in the browser. The admission list is never uploaded to any server.

## How the operator uses it

1. **Admission list:** upload the latest admission list (.xlsx). Only students marked Active are used.
2. **Test date & shifts:** pick the date and choose one shift or two, with the time of each. With two shifts, the same students keep the same seat, and there is one attendance sheet per shift.
3. **Batches & mode:** tick the batches appearing and set each one to Offline (classroom) or Online (computer lab). A batch with no known building asks for one. The choice is saved for next time.
4. **Rooms & capacity:** switch rooms on or off, raise capacities, and set the order in which students move between buildings.
5. **Generate sheet:** see the seat-fill map, reshuffle if wanted, and download the Excel file.

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
