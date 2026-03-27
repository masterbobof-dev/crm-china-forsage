export const npApi = async (apiKey: string, modelName: string, calledMethod: string, methodProperties: any = {}) => {
  if (!apiKey) throw new Error('API ключ Нової Пошти не вказано');
  const response = await fetch('https://api.novaposhta.ua/v2.0/json/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      apiKey,
      modelName,
      calledMethod,
      methodProperties
    })
  });
  const data = await response.json();
  if (!data.success) throw new Error(data.errors?.join(', ') || 'Помилка API Нової Пошти');
  return data.data;
};

export const trackParcel = (apiKey: string, documentNumber: string, phone: string = '') =>
  npApi(apiKey, 'TrackingDocument', 'getStatusDocuments', {
    Documents: [{ DocumentNumber: documentNumber, Phone: phone }]
  });

export const searchCities = (apiKey: string, cityName: string) => 
  npApi(apiKey, 'Address', 'getCities', { FindByString: cityName });

export const getWarehouses = (apiKey: string, cityRef: string) => 
  npApi(apiKey, 'Address', 'getWarehouses', { CityRef: cityRef });

export const getDocumentList = (apiKey: string, startDate: string, endDate: string) =>
  npApi(apiKey, 'InternetDocument', 'getDocumentList', {
    DateTimeFrom: startDate,
    DateTimeTo: endDate
  });
