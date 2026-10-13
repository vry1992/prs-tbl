export const dmsToDecimal = (coordinate) => {
  const value = String(coordinate).padStart(6, '0');

  const degrees = Number(value.slice(0, value.length - 4));
  const minutes = Number(value.slice(-4, -2));
  const seconds = Number(value.slice(-2));

  return degrees + minutes / 60 + seconds / 3600;
};
