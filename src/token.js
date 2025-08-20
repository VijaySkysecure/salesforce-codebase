async function refreshToken (refreshToken) {
    try {
        
    }
    catch (error) {
      console.error("Error refreshing token:", error.message);
      return res.status(500).send("Failed to refresh token");
    }
}