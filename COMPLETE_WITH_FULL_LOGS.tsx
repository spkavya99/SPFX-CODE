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
    console.log('Current URL:', window.location.href);
    
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
    } else {
      console.log('Not a Kaizen page, skipping image load');
    }
  }

  private isValidKaizenPage(): boolean {
    const url = window.location.href;
    const match = url.match(/Kaizen(\d+)/);
    const isValid = match !== null && match[1] !== undefined;
    console.log('isValidKaizenPage - URL:', url, 'Match:', match, 'IsValid:', isValid);
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
    console.log('Library Path:', libraryPath);
    console.log('Kaizen ID:', kaizenID);

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
        console.log('API Response Data:', data);
        console.log('Files in library:', data.value ? data.value.length : 0);
        
        if (data.value && data.value.length > 0) {
          console.log('All files in library:', data.value.map((f: any) => f.Name));
          
          const kaizenImages = data.value.filter((file: { Name: string }) => 
            file.Name.startsWith(kaizenID + '_')
          );
          
          console.log('Kaizen images found:', kaizenImages.length);
          console.log('Kaizen images:', kaizenImages.map((f: any) => f.Name));
          
          if (kaizenImages.length > 0) {
            const file = kaizenImages[0];
            const imageUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/${libraryPath}/${file.Name}?t=${Date.now()}`;
            
            console.log('Found user image:', file.Name);
            console.log('Image URL:', imageUrl);
            
            this.setState({
              imageUrl: imageUrl,
              imageName: file.Name,
              hasUserImage: true
            });
          } else {
            console.log('No Kaizen images found, showing default avatar');
            this.setState({
              imageUrl: this.defaultAvatarUrl,
              hasUserImage: false
            });
          }
        } else {
          console.log('Library is empty, showing default avatar');
          this.setState({
            imageUrl: this.defaultAvatarUrl,
            hasUserImage: false
          });
        }
      } else {
        console.error('API response not OK, status:', httpResponse.status);
        const errorText = await httpResponse.text();
        console.error('Error text:', errorText);
      }
    } catch (error: unknown) {
      console.error('=== loadImage ERROR ===', error);
      this.setState({
        imageUrl: this.defaultAvatarUrl,
        hasUserImage: false
      });
    }
    console.log('=== loadImage END ===');
  }

  private handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>): void => {
    console.log('=== handleFileSelect ===');
    const files: FileList | null = event.target.files;
    console.log('Files selected:', files ? files.length : 0);
    
    if (files && files.length > 0) {
      const file: File = files[0];
      console.log('File name:', file.name);
      console.log('File type:', file.type);
      console.log('File size:', file.size, 'bytes');
      
      if (!file.type.startsWith('image/')) {
        console.error('Not an image file');
        this.setState({ error: '❌ Please select an image file (jpg, png, gif, etc.)' });
        return;
      }
      
      if (file.size > 10 * 1024 * 1024) {
        console.error('File too large');
        this.setState({ error: '❌ Image size must be less than 10MB' });
        return;
      }

      console.log('File validation passed, starting upload');
      void this.uploadImage(file).catch((error: Error) => {
        console.error('Upload error caught:', error);
        this.setState({ error: '❌ ' + error.message });
      });
    }
  }

  private uploadImage = async (file: File): Promise<void> => {
    console.log('=== uploadImage START ===');
    console.log('File to upload:', file.name, 'Size:', file.size);
    
    this.setState({ uploading: true, message: 'Uploading...', error: '' });

    const libraryPath: string = this.getLibraryPath();
    const kaizenID: string = this.getKaizenID();
    
    const fileExtension: string = file.name.substring(file.name.lastIndexOf('.'));
    const newFileName: string = `${kaizenID}_image${fileExtension}`;
    
    console.log('New file name will be:', newFileName);
    console.log('Library path:', libraryPath);
    
    try {
      // Step 1: Delete existing file
      console.log('--- STEP 1: Delete old file (if exists) ---');
      const encodedFileName: string = encodeURIComponent(newFileName);
      const deleteUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${libraryPath}')/Files('${encodedFileName}')`;
      
      console.log('Delete URL:', deleteUrl);
      
      try {
        const deleteResponse = await this.props.context.spHttpClient.post(
          deleteUrl,
          SPHttpClient.configurations.v1,
          {
            headers: { 'X-HTTP-Method': 'DELETE' }
          }
        );
        console.log('Delete response status:', deleteResponse.status);
        console.log('Old file deleted successfully');
      } catch (deleteError: unknown) {
        console.log('No old file to delete or delete failed (this is OK):', deleteError);
      }

      // Step 2: Wait before uploading
      console.log('--- STEP 2: Waiting 500ms before upload ---');
      await new Promise(resolve => setTimeout(resolve, 500));

      // Step 3: Upload new file
      console.log('--- STEP 3: Upload new file ---');
      const uploadUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${libraryPath}')/Files/add(url='${encodedFileName}',overwrite=false)`;
      
      console.log('Upload URL:', uploadUrl);
      
      const arrayBuffer = await this.readFileAsArrayBuffer(file);
      console.log('File converted to ArrayBuffer, size:', arrayBuffer.byteLength, 'bytes');

      console.log('Sending POST request to upload...');
      const httpResponse: SPHttpClientResponse = await this.props.context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          body: arrayBuffer
        }
      );

      console.log('=== UPLOAD RESPONSE ===');
      console.log('Status:', httpResponse.status);
      console.log('Status text:', httpResponse.statusText);
      console.log('OK:', httpResponse.ok);

      if (httpResponse.ok) {
        console.log('✅ Upload successful!');
        const responseData = await httpResponse.json();
        console.log('Response data:', responseData);
        
        this.setState({ 
          message: '✅ Image uploaded successfully!',
          uploading: false,
          error: ''
        });
        
        setTimeout(() => {
          console.log('Reloading image after upload...');
          void this.loadImage().catch((error: Error) => {
            console.error('Error reloading image:', error);
          });
        }, 2000);
        
        setTimeout(() => {
          this.setState({ message: '' });
        }, 4000);
      } else {
        console.error('❌ Upload failed!');
        console.error('Status:', httpResponse.status);
        
        const responseText = await httpResponse.text();
        console.error('Response text:', responseText);
        
        try {
          const errorJson = JSON.parse(responseText);
          console.error('Error JSON:', errorJson);
        } catch (e) {
          console.error('Could not parse error as JSON');
        }
        
        this.setState({ 
          message: '',
          error: '❌ Upload failed. Status: ' + httpResponse.status + '. Check console for details.',
          uploading: false 
        });
      }
    } catch (error: unknown) {
      console.error('=== UPLOAD EXCEPTION ===');
      console.error('Error type:', typeof error);
      console.error('Error:', error);
      
      if (error instanceof Error) {
        console.error('Error message:', error.message);
        console.error('Error stack:', error.stack);
      }
      
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.setState({ 
        message: '',
        error: '❌ Error uploading image: ' + errorMessage,
        uploading: false 
      });
    }

    if (this.fileInput.current) {
      this.fileInput.current.value = '';
    }
    
    console.log('=== uploadImage END ===');
  }

  private readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    console.log('--- readFileAsArrayBuffer START ---');
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = () => {
        console.log('FileReader onload called');
        if (reader.result instanceof ArrayBuffer) {
          console.log('Successfully read file as ArrayBuffer');
          resolve(reader.result);
        } else {
          console.error('Reader result is not ArrayBuffer');
          reject(new Error('Failed to read file as ArrayBuffer'));
        }
      };
      
      reader.onerror = () => {
        console.error('FileReader error:', reader.error);
        reject(new Error('Error reading file: ' + reader.error));
      };
      
      reader.onprogress = (event) => {
        console.log('FileReader progress:', event.loaded, 'of', event.total);
      };
      
      console.log('Starting FileReader.readAsArrayBuffer');
      reader.readAsArrayBuffer(file);
    });
  }

  private deleteImage = async (): Promise<void> => {
    console.log('=== deleteImage START ===');
    if (!window.confirm('Are you sure you want to delete this image and go back to default?')) {
      console.log('Delete cancelled by user');
      return;
    }

    const libraryPath: string = this.getLibraryPath();
    const imageName: string = this.state.imageName;
    console.log('Deleting image:', imageName);
    
    const encodedFileName: string = encodeURIComponent(imageName);
    const deleteUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${libraryPath}')/Files('${encodedFileName}')`;

    console.log('Delete URL:', deleteUrl);

    try {
      this.setState({ uploading: true });
      
      const httpResponse = await this.props.context.spHttpClient.post(
        deleteUrl,
        SPHttpClient.configurations.v1,
        {
          headers: { 'X-HTTP-Method': 'DELETE' }
        }
      );
      
      console.log('Delete response status:', httpResponse.status);
      
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
    console.log('=== deleteImage END ===');
  }

  public render(): React.ReactElement<IImageUploadWebPartProps> {
    console.log('=== render ===');
    console.log('State:', this.state);
    
    const { imageUrl, uploading, message, error, isKaizenPage, hasUserImage } = this.state;

    if (!isKaizenPage) {
      console.log('Not rendering image upload - not a Kaizen page');
      return (
        <div className={styles.imageUpload}>
          <p style={{ color: '#666', fontStyle: 'italic', textAlign: 'center', padding: '20px' }}>
            This web part only works on individual Kaizen pages, not on the template page.
          </p>
        </div>
      );
    }

    console.log('Rendering image upload interface');
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
                console.log('Image failed to load:', imageUrl);
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
                    console.log('Delete button clicked');
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
