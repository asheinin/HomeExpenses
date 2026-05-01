process.env.TZ = 'America/New_York';
var d = new Date("2026-01-01T00:00:00");
var endDate = new Date("2026-12-01T00:00:00");
while(d <= endDate) {
  console.log(d.toISOString(), d.getTime());
  d.setMonth(d.getMonth() + 1);
}
