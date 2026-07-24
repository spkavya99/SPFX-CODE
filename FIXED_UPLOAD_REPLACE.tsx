import * as React from 'react';
import styles from './ImageUploadWebPart.module.scss';
import { IImageUploadWebPartProps } from './IImageUploadWebPartProps';
import { SPHttpClient } from '@microsoft/sp-http';

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
    
    const isKaizenPage = this.isValidKaizenPage();
    
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
    if (this.state.isKaizenPage) {
      void this.loadImage().catch((error: Error) => {
        console.error('Error loading image on mount:', error);
      });
    }
  }

  private isValidKaizenPage(): boolean {
    const url = window.location.href;
    const match = url.match(/Kaizen(\d+)/);
    return match !== null && match[1] !== undefined;
  }

  private getKaizenID(): string {
    const url = window.location.href;
    const match = url.match(/Kaizen(\d+)/);
    return match ? match[1] : 'default';
  }

  private getLibraryPath(): string {
    return `KaizenImages`;
  }

  private loadImage = async (): Promise<void> => {
    const libraryPath: string = this.getLibraryPath();
    const kaizenID: string = this.getKaizenID();

    try {
      const url: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${libraryPath}')/Files`;
      
      const httpResponse = await this.props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (httpResponse.ok) {
        const data = await httpResponse.json();
        
        if (data.value && data.value.length > 0) {
          const kaizenImages = data.value.filter((file: { Name: string }) => 
            file.Name.startsWith(kaizenID + '_')
          );
          
          if (kaizenImages.length > 0) {
            const file = kaizenImages[0];
            const imageUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/${libraryPath}/${file.Name}?t=${Date.now()}`;
            
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
      console.log('No images found or library not accessible, showing default avatar');
      this.setState({
        imageUrl: this.defaultAvatarUrl,
        hasUserImage: false
      });
    }
  }

  private handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files: FileList | null = event.target.files;
    if (files && files.length > 0) {
      const file: File = files[0];
      
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
      });
    }
  }

  private uploadImage = async (file: File): Promise<void> => {
    this.setState({ uploading: true, message: 'Uploading...', error: '' });

    const libraryPath: string = this.getLibraryPath();
    const kaizenID: string = this.getKaizenID();
    
    const fileExtension: string = file.name.substring(file.name.lastIndexOf('.'));
    const newFileName: string = `${kaizenID}_image${fileExtension}`;
    
    const encodedFileName: string = encodeURIComponent(newFileName);
    const uploadUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${libraryPath}')/Files/add(url='${encodedFileName}',overwrite=true)`;

    try {
      // Read file as ArrayBuffer
      const arrayBuffer = await this.readFileAsArrayBuffer(file);
      
      const httpResponse = await this.props.context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          body: arrayBuffer
        }
      );

      if (httpResponse.ok) {
        this.setState({ 
          message: '✅ Image uploaded successfully!',
          uploading: false,
          error: ''
        });
        
        setTimeout(() => {
          void this.loadImage().catch((error: Error) => {
            console.error('Error reloading image:', error);
          });
        }, 2000);
        
        setTimeout(() => {
          this.setState({ message: '' });
        }, 4000);
      } else {
        const errorText = await httpResponse.text();
        console.error('Upload error response:', errorText);
        this.setState({ 
          message: '',
          error: '❌ Upload failed. Please try again.',
          uploading: false 
        });
      }
    } catch (error: unknown) {
      console.error('Upload error:', error);
      this.setState({ 
        message: '',
        error: '❌ Error uploading image. ' + (error instanceof Error ? error.message : ''),
        uploading: false 
      });
    }

    if (this.fileInput.current) {
      this.fileInput.current.value = '';
    }
  }

  private readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (reader.result instanceof ArrayBuffer) {
          resolve(reader.result);
        } else {
          reject(new Error('Failed to read file as ArrayBuffer'));
        }
      };
      reader.onerror = () => {
        reject(new Error('Error reading file'));
      };
      reader.readAsArrayBuffer(file);
    });
  }

  private deleteImage = async (): Promise<void> => {
    if (!window.confirm('Are you sure you want to delete this image and go back to default?')) {
      return;
    }

    const libraryPath: string = this.getLibraryPath();
    const imageName: string = this.state.imageName;
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
                onClick={() => this.fileInput.current?.click()}
                disabled={uploading}
              >
                🔄 Replace
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
