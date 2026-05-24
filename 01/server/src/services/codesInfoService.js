import axiosInstance from "../utils/axios.js";

export const getCodesInfo = async (codes, pg) => {
  try {
    const { data } = await axiosInstance.post(
      `/cises/publicInfo?pg=${pg}`,
      codes
    );

    return data;
  } catch (error) {
    console.error("Error:", error);
  }
};
