import * as React from "react";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { ServiceScope } from "@microsoft/sp-core-library";

const LIVE_PERSONA_COMPONENT_ID = "914330ee-2df2-4f6e-a858-30c23a812408";

export interface ILivePersonaCardProps {
  serviceScope: ServiceScope;
  upn: string;
  template: string | JSX.Element;
}

export const LivePersonaCard: React.FC<ILivePersonaCardProps> =
  function LivePersonaCard(
    props: React.PropsWithChildren<ILivePersonaCardProps>
  ) {
    const [loading, setLoading] = React.useState(true);
    const sharedLibrary = React.useRef<any>();
    const { serviceScope, upn, template } = props;

    React.useEffect(() => {
      const fetchComponent = async (): Promise<void> => {
        if (loading) {
          try {
            sharedLibrary.current = await SPComponentLoader.loadComponentById(
              LIVE_PERSONA_COMPONENT_ID
            );
          } catch (error) {
            console.error(`[LivePersona] ${error}`);
          }
          setLoading(false);
        }
      };

      fetchComponent().catch(() => {
        /* no-op; */
      });
    }, []);

    let renderPersona: JSX.Element = React.createElement("div", {}, template);

    if (loading === false) {
      renderPersona = React.createElement(
        sharedLibrary.current.LivePersonaCard,
        {
          className: "livePersonaCard",
          clientScenario: "livePersonaCard",
          disableHover: false,
          hostAppPersonaInfo: {
            PersonaType: "User",
          },
          upn: upn,
          legacyUpn: upn,
          serviceScope: serviceScope,
        },
        React.createElement("div", {}, template)
      );
    }

    return renderPersona;
  };
