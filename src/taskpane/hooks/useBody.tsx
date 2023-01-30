import { useState, useEffect } from "react";

const useBody = () => {
  const [body, setBody] = useState("");

  const getBody = () => {
    Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      setBody(body.text);
    });
  };

  useEffect(() => {
    setInterval(getBody, 100);
  }, []);

  return body;
};

export default useBody;
