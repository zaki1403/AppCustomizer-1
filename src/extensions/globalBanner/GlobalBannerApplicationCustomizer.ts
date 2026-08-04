import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

const LOG_SOURCE: string = 'GlobalBannerApplicationCustomizer';

export interface IGlobalBannerApplicationCustomizerProperties {
  topMessage?: string;
  bottomMessage?: string;
}

export default class GlobalBannerApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalBannerApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initializing Header...`);

    // Listen for placeholder changes
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholders);

    return Promise.resolve();
  }

  private _renderPlaceholders(): void {
    // Relative path to your newly uploaded image: sp3.jpg
    const imageUrl = '/SiteAssets/SitePages/Home/sp6.jpg';

    // Render Top Header
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (this._topPlaceholder && this._topPlaceholder.domElement) {
        // Get current user's display name
        const userName = this.context.pageContext.user.displayName || 'User';

        // Personalized welcome message
        const topMessage = `Welcome To CrystaliX ${userName}`;
        
        const headerStyles = `
          background-image: url('${imageUrl}'); 
          background-size: cover; 
          background-position: center; 
          height: 100px; 
          display: flex; 
          align-items: center; 
          justify-content: center; 
          font-family: 'Segoe UI', 'Helvetica Neue', Helvetica, sans-serif;
        `;

        const headerTextStyles = `
          color: #ffffff; 
          font-weight: 600; 
          font-size: 22px; 
          padding: 0 20px;
          text-shadow: 0px 2px 5px rgba(0,0,0,0.8); 
          text-align: center;
        `;

        this._topPlaceholder.domElement.innerHTML = `
          <div style="${headerStyles}">
            <span style="${headerTextStyles}">${topMessage}</span>
          </div>`;
      }
    }
  }

  private _onDispose(): void {
    console.log('[GlobalBannerApplicationCustomizer] Disposed placeholders.');
  }
}
