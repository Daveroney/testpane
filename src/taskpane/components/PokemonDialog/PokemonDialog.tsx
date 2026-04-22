import React, { useEffect } from "react";

export function PokemonDialog() {
  useEffect(() => {
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      async (arg: Office.DialogParentMessageReceivedEventArgs) => {
        const message = JSON.parse(arg.message);
        if (message.type === "filteredPartners") {
          getPartners(message);
        }
        if (message.type === "documentData") {
          // Get access to the current draft email
          setId(message.id);
          setGraphAccessToken(message.accessToken);
          setResourceUrl(message.resourceUrl);
          setDocumentType(message.documentType);
        }
      }
    );
    Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
  }, []);
  return <div>PokemonDialog</div>;
}
