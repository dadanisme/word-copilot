import React, { useEffect } from "react";
import { DefaultButton } from "@fluentui/react";
import { useBody, useDebounce } from "../hooks";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App = (props: AppProps) => {
  const body = useBody();
  const debouncedBody = useDebounce(body, 1000);

  useEffect(() => {
    if (debouncedBody) {
      click();
    }
  }, [debouncedBody]);

  const getResponse = async (prompt: string): Promise<string> => {
    console.log("getting response");

    // const aiPrompt = "complete the following sentence: \n" + prompt;
    const res = await fetch("http://localhost:8000/chatgpt/" + prompt);
    const data = (await res.json()) as Record<string, { text: string }[]>;

    return data.choices[0].text;
  };

  const click = async () => {
    return Word.run(async (context) => {
      console.log({ context });

      const pList = context.document.body.paragraphs;

      const data = [];

      pList.load("text");
      await context.sync();

      for (let i = 0; i < pList.items.length; i++) {
        data.push(pList.items[i].text);
      }

      const last = data[data.length - 1];

      const response = await getResponse(last);

      // split response into sentences, delimited by double newlines
      const sentences = response.split("\n\n");

      // insert each sentence into the document
      for (let i = 0; i < sentences.length; i++) {
        const paragraph = context.document.body.insertText(sentences[i], Word.InsertLocation.end);
        paragraph.font.color = "blue";
        paragraph.font.italic = true;

        // insert a newline after each sentence
        context.document.body.insertParagraph("", Word.InsertLocation.end);
      }

      await context.sync();

      // add line break after the last paragraph
      context.document.body.insertParagraph("", Word.InsertLocation.end);

      // reset styles
      context.document.body.font.color = "black";
      context.document.body.font.italic = false;
    });
  };

  const { isOfficeInitialized } = props;

  if (!isOfficeInitialized) {
    return <div>Loading...</div>;
  }

  return (
    <div className="ms-welcome">
      <p className="ms-font-l">
        Modify the source files, then click <b>Run</b>.
      </p>
      <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
        Run
      </DefaultButton>
    </div>
  );
};

export default App;
