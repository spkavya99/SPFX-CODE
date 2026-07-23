// File: src/webparts/imageUploadWebPart/components/ImageUploadWebPart.tsx

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
}

export default class ImageUploadWebPart extends React.Component<IImageUploadWebPartProps, IState> {
  private fileInput: React.RefObject<HTMLInputElement>;

  constructor(props: IImageUploadWebPartProps) {
    super(props);
    this.fileInput = React.createRef();
    this.state = {
      imageUrl: null,
      imageName: '',
      uploading: false,
      message: '',
      error: ''
    };
  }

  componentDidMount(): void {
    this.loadImage();
  }

  private getKaizenID(): string {
    const url = window.location.href;
    const match = url.match(/Kaizen(\d+)/);
    return match ? match[1] : 'default';
  }

  private getLibraryPath(): string {
    return `/sites/NXT-CBS/KaizenImages`;
  }

  private loadImage = async (): Promise<void> => {
    const libraryPath = this.getLibraryPath();
    const kaizenID = this.getKaizenID();

    try {
      const url: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryPath}')/Files`;
      const response: any = await this.props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data: any = await response.json();
        
        if (data.value && data.value.length > 0) {
          const kaizenImages = data.value.filter((file: any) => 
            file.Name.startsWith(kaizenID + '_')
          );
          
          if (kaizenImages.length > 0) {
            const file = kaizenImages[0];
            const imageUrl: string = `${this.props.context.pageContext.web.absoluteUrl}${file.ServerRelativeUrl}`;
            
            this.setState({
              imageUrl: imageUrl,
              imageName: file.Name
            });
          }
        }
      }
    } catch (error: any) {
      console.log('No images found or library not accessible');
    }
  }

  private handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const file = files[0];
      
      // Validate file type
      if (!file.type.startsWith('image/')) {
        this.setState({ error: '❌ Please select an image file (jpg, png, gif, etc.)' });
        return;
      }
      
      // Validate file size (max 10MB)
      if (file.size > 10 * 1024 * 1024) {
        this.setState({ error: '❌ Image size must be less than 10MB' });
        return;
      }

      this.uploadImage(file);
    }
  }

  private uploadImage = async (file: File): Promise<void> => {
    this.setState({ uploading: true, message: 'Uploading...', error: '' });

    const libraryPath: string = this.getLibraryPath();
    const kaizenID: string = this.getKaizenID();
    
    // Create filename: [KaizenID]_image.[extension]
    const fileExtension: string = file.name.substring(file.name.lastIndexOf('.'));
    const newFileName: string = `${kaizenID}_image${fileExtension}`;
    
    const encodedFileName: string = encodeURIComponent(newFileName);
    const uploadUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryPath}')/Files/add(url='${encodedFileName}',overwrite=true)`;

    try {
      const response: any = await this.props.context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          body: file
        }
      );

      if (response.ok) {
        this.setState({ 
          message: '✅ Image uploaded successfully!',
          uploading: false,
          error: ''
        });
        
        // Reload image after upload
        setTimeout(() => {
          this.loadImage();
        }, 500);
        
        // Clear message after 3 seconds
        setTimeout(() => {
          this.setState({ message: '' });
        }, 3000);
      } else {
        this.setState({ 
          message: '',
          error: '❌ Upload failed. Please try again.',
          uploading: false 
        });
      }
    } catch (error: any) {
      console.error('Upload error:', error);
      this.setState({ 
        message: '',
        error: '❌ Error uploading image.',
        uploading: false 
      });
    }

    // Reset file input
    if (this.fileInput.current) {
      this.fileInput.current.value = '';
    }
  }

  private deleteImage = async (): Promise<void> => {
    if (!window.confirm('Are you sure you want to delete this image?')) {
      return;
    }

    const libraryPath: string = this.getLibraryPath();
    const imageName: string = this.state.imageName;
    const encodedFileName: string = encodeURIComponent(imageName);
    
    const deleteUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryPath}')/Files('${encodedFileName}')`;

    try {
      this.setState({ uploading: true });
      
      const response: any = await this.props.context.spHttpClient.post(
        deleteUrl,
        SPHttpClient.configurations.v1,
        {
          headers: { 'X-HTTP-Method': 'DELETE' }
        }
      );
      
      this.setState({ 
        imageUrl: null, 
        imageName: '',
        uploading: false,
        message: '✅ Image deleted!'
      });

      setTimeout(() => {
        this.setState({ message: '' });
      }, 2000);
    } catch (error: any) {
      console.error('Delete error:', error);
      this.setState({ 
        error: '❌ Error deleting image',
        uploading: false 
      });
    }
  }

  public render(): React.ReactElement<IImageUploadWebPartProps> {
    const { imageUrl, uploading, message, error } = this.state;

    return (
      <div className={styles.imageUpload}>
        <h3>📸 Image</h3>
        
        {message && <p className={styles.message}>{message}</p>}
        {error && <p className={styles.error}>{error}</p>}

        <button 
          className={styles.uploadButton}
          onClick={() => this.fileInput.current?.click()}
          disabled={uploading}
        >
          {uploading ? 'Uploading...' : '+ Add Image'}
        </button>

        <input
          ref={this.fileInput}
          type="file"
          accept="image/*"
          onChange={this.handleFileSelect}
          style={{ display: 'none' }}
        />

        {imageUrl && (
          <div className={styles.imageContainer}>
            <img 
              src={imageUrl} 
              alt="Kaizen" 
              className={styles.singleImage}
              onError={() => {
                this.setState({ error: '❌ Could not load image' });
              }}
            />
            <div className={styles.buttonGroup}>
              <button 
                className={styles.deleteButton}
                onClick={this.deleteImage}
                disabled={uploading}
              >
                🗑️ Delete
              </button>
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

        {!imageUrl && !message && (
          <p className={styles.placeholder}>No image uploaded yet. Click "+ Add Image" to upload.</p>
        )}
      </div>
    );
  }
}
