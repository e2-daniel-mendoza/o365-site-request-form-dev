import {
    ThemeProvider,
    IReadonlyTheme,
    ThemeChangedEventArgs
} from '@microsoft/sp-component-base';
import { WebPartContext, BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import SiteRequestFormWebPart from '../SiteRequestFormWebPart';

export default class ThemeVariantService {

    private static _themeProvider: ThemeProvider;

    public static themeVariant: IReadonlyTheme | undefined;

    public static Initialize(initWebPart: SiteRequestFormWebPart): void {
        this.themeVariant = initWebPart.context.serviceScope.consume(ThemeProvider.serviceKey).tryGetTheme();
        initWebPart.context.serviceScope.consume(ThemeProvider.serviceKey).themeChangedEvent.add(initWebPart, result => {
            console.log(result.theme.semanticColors.bodyBackground);
            this.themeVariant = result.theme;
            initWebPart.render();
        });
    }
}