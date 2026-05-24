import axiosInstance from "../utils/axios.js";

export const getCodesHistory = async (id) => {
  try {
    const encodedId = encodeURIComponent(id);
    const { data } = await axiosInstance.get(
      `/cises/history?code=${encodedId}`
    );
    return data;
  } catch (error) {
    console.error("Error:", error);
    return { cis: id, error: "Not Found!" };
  }
};
