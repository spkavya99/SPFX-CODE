import * as React from 'react';
import styles from './ImageUploadWebPart.module.scss';
import { IImageUploadWebPartProps } from './IImageUploadWebPartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IState {
  imageUrl: string | null;
  imageName: string;
  uploading: boolean;
  message: string;
  error: string;
  isKaizenPage: boolean;
  hasUserImage: boolean;
}

export default class ImageUploadWebPart extends React.Component<IImageUploadWebPartProps, IState> {
  private fileInput: React.RefObject<HTMLInputElement>;
  private defaultAvatarUrl: string = 'https://cranenxt.sharepoint.com/sites/NXT-CBS/SiteAssets/default-avatar.png';

  constructor(props: IImageUploadWebPartProps) {
    super(props);
    this.fileInput = React.createRef();
    
    console.log('=== ImageUploadWebPart Constructor ===');
    console.log('Web URL:', this.props.context.pageContext.web.absoluteUrl);
    
    const isKaizenPage = this.isValidKaizenPage();
    console.log('Is Kaizen Page:', isKaizenPage);
    
    this.state = {
      imageUrl: this.defaultAvatarUrl,
      imageName: '',
      uploading: false,
      message: '',
      error: '',
      isKaizenPage: isKaizenPage,
      hasUserImage: false
    };
  }

  public componentDidMount(): void {
    console.log('=== componentDidMount ===');
    if (this.state.isKaizenPage) {
      console.log('Loading image for Kaizen page');
      void this.loadImage().catch((error: Error) => {
        console.error('Error loading image on mount:', error);
      });
    }
  }

  private isValidKaizenPage(): boolean {
    const url = window.location.href;
    const match = url.match(/Kaizen(\d+)/);
    const isValid = match !== null && match[1] !== undefined;
    console.log('isValidKaizenPage:', isValid);
    return isValid;
  }

  private getKaizenID(): string {
    const url = window.location.href;
    const match = url.match(/Kaizen(\d+)/);
    const kaizenID = match ? match[1] : 'default';
    console.log('getKaizenID:', kaizenID);
    return kaizenID;
  }

  private getLibraryPath(): string {
    return `KaizenImages`;
  }

  private loadImage = async (): Promise<void> => {
    console.log('=== loadImage START ===');
    const libraryPath: string = this.getLibraryPath();
    const kaizenID: string = this.getKaizenID();

    try {
      const url: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${libraryPath}')/Files`;
      console.log('API URL:', url);
      
      const httpResponse = await this.props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      console.log('API Response Status:', httpResponse.status);

      if (httpResponse.ok) {
        const data = await httpResponse.json();
        console.log('Files in library:', data.value ? data.value.length : 0);
        
        if (data.value && data.value.length > 0) {
          const kaizenID: string = this.getKaizenID();
          const kaizenImages = data.value.filter((file: { Name: string }) => 
            file.Name.startsWith(kaizenID + '_')
          );
          
          console.log('Kaizen images found:', kaizenImages.length);
          
          if (kaizenImages.length > 0) {
            const file = kaizenImages[0];
            const imageUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/${libraryPath}/${file.Name}?t=${Date.now()}`;
            
            console.log('Found user image:', file.Name);
            
            this.setState({
              imageUrl: imageUrl,
              imageName: file.Name,
              hasUserImage: true
            });
          } else {
            this.setState({
              imageUrl: this.defaultAvatarUrl,
              hasUserImage: false
            });
          }
        } else {
          this.setState({
            imageUrl: this.defaultAvatarUrl,
            hasUserImage: false
          });
        }
      }
    } catch (error: unknown) {
      console.error('loadImage ERROR:', error);
      this.setState({
        imageUrl: this.defaultAvatarUrl,
        hasUserImage: false
      });
    }
  }

  private handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>): void => {
    console.log('=== handleFileSelect ===');
    const files: FileList | null = event.target.files;
    console.log('Files selected:', files ? files.length : 0);
    
    if (files && files.length > 0) {
      const file: File = files[0];
      console.log('File name:', file.name);
      console.log('File type:', file.type);
      console.log('File size:', file.size);
      
      if (!file.type.startsWith('image/')) {
        this.setState({ error: '❌ Please select an image file (jpg, png, gif, etc.)' });
        return;
      }
      
      if (file.size > 10 * 1024 * 1024) {
        this.setState({ error: '❌ Image size must be less than 10MB' });
        return;
      }

      void this.uploadImage(file).catch((error: Error) => {
        console.error('Upload error:', error);
        this.setState({ error: '❌ ' + error.message });
      });
    }
  }

  private uploadImage = async (file: File): Promise<void> => {
    console.log('=== uploadImage START ===');
    console.log('File:', file.name, 'Size:', file.size);
    
    this.setState({ uploading: true, message: 'Uploading...', error: '' });

    const libraryPath: string = this.getLibraryPath();
    const kaizenID: string = this.getKaizenID();
    const fileExtension: string = file.name.substring(file.name.lastIndexOf('.'));
    const newFileName: string = `${kaizenID}_image${fileExtension}`;
    
    console.log('New file name:', newFileName);
    
    try {
      // Step 1: Convert file to blob
      console.log('--- Step 1: Convert file to Blob ---');
      const blob = new Blob([file], { type: file.type });
      console.log('Blob created, size:', blob.size);

      // Step 2: Upload file
      console.log('--- Step 2: Upload file ---');
      const encodedFileName: string = encodeURIComponent(newFileName);
      const uploadUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${libraryPath}')/Files/add(url='${encodedFileName}',overwrite=true)`;
      
      console.log('Upload URL:', uploadUrl);
      console.log('Sending upload request...');

      const httpResponse: SPHttpClientResponse = await this.props.context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          body: blob
        }
      );

      console.log('=== UPLOAD RESPONSE ===');
      console.log('Status:', httpResponse.status);
      console.log('Status text:', httpResponse.statusText);
      console.log('OK:', httpResponse.ok);

      if (httpResponse.ok) {
        console.log('✅ Upload successful!');
        
        this.setState({ 
          message: '✅ Image uploaded successfully!',
          uploading: false,
          error: ''
        });
        
        setTimeout(() => {
          console.log('Reloading image...');
          void this.loadImage().catch((error: Error) => {
            console.error('Error reloading:', error);
          });
        }, 2000);
        
        setTimeout(() => {
          this.setState({ message: '' });
        }, 4000);
      } else {
        console.error('❌ Upload failed!');
        console.error('Status:', httpResponse.status);
        
        const responseText = await httpResponse.text();
        console.error('Response:', responseText);
        
        this.setState({ 
          message: '',
          error: '❌ Upload failed. Status: ' + httpResponse.status,
          uploading: false 
        });
      }
    } catch (error: unknown) {
      console.error('=== UPLOAD EXCEPTION ===');
      console.error('Error:', error);
      
      const errorMessage = error instanceof Error ? error.message : String(error);
      this.setState({ 
        message: '',
        error: '❌ Error: ' + errorMessage,
        uploading: false 
      });
    }

    if (this.fileInput.current) {
      this.fileInput.current.value = '';
    }
    
    console.log('=== uploadImage END ===');
  }

  private deleteImage = async (): Promise<void> => {
    console.log('=== deleteImage START ===');
    if (!window.confirm('Are you sure you want to delete this image and go back to default?')) {
      return;
    }

    const libraryPath: string = this.getLibraryPath();
    const imageName: string = this.state.imageName;
    console.log('Deleting:', imageName);
    
    const encodedFileName: string = encodeURIComponent(imageName);
    const deleteUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${libraryPath}')/Files('${encodedFileName}')`;

    try {
      this.setState({ uploading: true });
      
      const httpResponse = await this.props.context.spHttpClient.post(
        deleteUrl,
        SPHttpClient.configurations.v1,
        {
          headers: { 'X-HTTP-Method': 'DELETE' }
        }
      );
      
      console.log('Delete status:', httpResponse.status);
      
      this.setState({ 
        imageUrl: this.defaultAvatarUrl,
        imageName: '',
        uploading: false,
        hasUserImage: false,
        message: '✅ Image deleted! Back to default avatar.'
      });

      setTimeout(() => {
        this.setState({ message: '' });
      }, 2000);
    } catch (error: unknown) {
      console.error('Delete error:', error);
      this.setState({ 
        error: '❌ Error deleting image',
        uploading: false 
      });
    }
  }

  public render(): React.ReactElement<IImageUploadWebPartProps> {
    const { imageUrl, uploading, message, error, isKaizenPage, hasUserImage } = this.state;

    if (!isKaizenPage) {
      return (
        <div className={styles.imageUpload}>
          <p style={{ color: '#666', fontStyle: 'italic', textAlign: 'center', padding: '20px' }}>
            This web part only works on individual Kaizen pages, not on the template page.
          </p>
        </div>
      );
    }

    return (
      <div className={styles.imageUpload}>
        <h3>📸 Image</h3>
        
        {message && <p className={styles.message}>{message}</p>}
        {error && <p className={styles.error}>{error}</p>}

        {imageUrl && (
          <div className={styles.imageContainer}>
            <img 
              src={imageUrl} 
              alt="Kaizen" 
              className={styles.singleImage}
              onError={() => {
                console.log('Image failed to load');
                this.setState({ 
                  imageUrl: this.defaultAvatarUrl,
                  error: '❌ Could not load image. Showing default avatar.' 
                });
              }}
            />
            <div className={styles.buttonGroup}>
              {hasUserImage && (
                <button 
                  className={styles.deleteButton}
                  onClick={() => {
                    void this.deleteImage().catch((error: Error) => {
                      console.error('Delete error:', error);
                    });
                  }}
                  disabled={uploading}
                >
                  🗑️ Delete
                </button>
              )}
              <button 
                className={styles.replaceButton}
                onClick={() => {
                  console.log('Replace button clicked');
                  this.fileInput.current?.click();
                }}
                disabled={uploading}
              >
                {uploading ? 'Uploading...' : '🔄 Replace'}
              </button>
            </div>
          </div>
        )}

        <input
          ref={this.fileInput}
          type="file"
          accept="image/*"
          onChange={this.handleFileSelect}
          style={{ display: 'none' }}
        />
      </div>
    );
  }
}
