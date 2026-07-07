window.causalesData = [];
window.causalesLoadPromise = fetch('./causales.json')
  .then((response) => {
    if (!response.ok) {
      throw new Error('No se pudo cargar causales.json');
    }
    return response.json();
  })
  .then((data) => {
    window.causalesData = Array.isArray(data) ? data : [];
    return window.causalesData;
  })
  .catch(() => window.causalesData);
