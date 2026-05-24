import axios from "axios";

const BASE_URL = "https://goods.aslbelgisi.uz/api/v3/true-api";
let authToken = "токен от ЭЦП";

const axiosInstance = axios.create({
  baseURL: BASE_URL,
});
axiosInstance.interceptors.request.use(
  async (config) => {
    config.headers.Authorization = `Bearer ${authToken}`;
    return config;
  },
  (error) => Promise.reject(error)
);

export default axiosInstance;
