document.getElementById("buka").addEventListener("click", () => {
  chrome.tabs.create({ url: "https://studentportal.ipb.ac.id/Kegiatan/AktivitasKampusMerdeka" });
  window.close();
});
