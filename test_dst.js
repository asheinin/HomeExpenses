process.env.TZ = 'America/New_York';
const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
var startDate = new Date("2026-01-01T00:00:00");
var endDate = new Date("2026-12-01T00:00:00");
const targetSheets = [];
let d = new Date(startDate.getTime());
while (d <= endDate) {
  targetSheets.push({ monthStr: months[d.getMonth()], year: d.getFullYear() });
  d.setMonth(d.getMonth() + 1);
}
console.log(targetSheets);
