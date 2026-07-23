# SPFx Image Upload - Single KaizenImages Library

## ✅ SIMPLIFIED APPROACH

All images upload to **ONE central library** called `KaizenImages`

```
KaizenImages Library:
├── 9951_photo.jpg
├── 9950_image.jpg  
├── 9949_screenshot.jpg
└── ...

Each filename includes KaizenID for easy identification
```

---

# PART 1: SETUP & PREREQUISITES

## Step 1: Install Required Software

### 1.1 Install Node.js
- Download from: https://nodejs.org/
- Choose **LTS version**
- Install normally
- Verify:
```bash
node --version
npm --version
```

### 1.2 Install Global Tools
```bash
npm install -g @microsoft/sharepoint@latest
npm install -g yo
npm install -g @microsoft/generator-sharepoint
```

---

# PART 2: CREATE SPFx PROJECT

## Step 2: Create Project Directory
```bash
mkdir ImageUploadWebPart
cd ImageUploadWebPart
```

## Step 3: Generate SPFx Project
```bash
yo @microsoft/sharepoint
```

**Answer prompts:**
```
? New SharePoint client-side web part in a new solution? Yes
? What's your solution name? ImageUploadWebPart
? Which type of client-side component to create? WebPart
? What's your Web part name? ImageUploadWebPart
? Which template would you like to use? React
? Does the solution include the Microsoft Graph SDK? No
? Would you like to use PnP SPFx Scaffolding? No
```

## Step 4: Install Dependencies
```bash
npm install @microsoft/sp-http
npm install uuid
```

---

# PART 3: REPLACE CODE FILES

## Step 5: ImageUploadWebPart.tsx

**Path:** `src/webparts/imageUploadWebPart/components/ImageUploadWebPart.tsx`

**Replace with:**

```typescript
import * as React from 'react';
import styles from './ImageUploadWebPart.module.scss';
import { IImageUploadWebPartProps } from './IImageUploadWebPartProps';
import { SPHttpClient } from '@microsoft/sp-http';

export default class ImageUploadWebPart extends React.Component<IImageUploadWebPartProps, any> {
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

  componentDidMount() {
    this.loadImage();
  }

  // Get KaizenID from page URL (e.g., Kaizen9951)
  getKaizenID = (): string => {
    const url = window.location.href;
    const match = url.match(/Kaizen(\d+)/);
    return match ? match[1] : 'default';
  }

  // Get KaizenImages library path
  getLibraryPath = (): string => {
    return `/sites/NXT-CBS/KaizenImages`;
  }

  // Load existing image for this Kaizen from KaizenImages library
  loadImage = async () => {
    const libraryPath = this.getLibraryPath();
    const kaizenID = this.getKaizenID();

    try {
      const response = await this.props.context.spHttpClient.get(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryPath}')/Files`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        
        // Find image for this Kaizen (starts with KaizenID)
        if (data.value && data.value.length > 0) {
          const kaizenImages = data.value.filter((file: any) => 
            file.Name.startsWith(kaizenID + '_')
          );
          
          if (kaizenImages.length > 0) {
            const file = kaizenImages[0];
            const imageUrl = `${this.props.context.pageContext.web.absoluteUrl}${file.ServerRelativeUrl}`;
            
            this.setState({
              imageUrl: imageUrl,
              imageName: file.Name
            });
          }
        }
      }
    } catch (error) {
      console.log('Library not found or no images yet');
    }
  }

  // Handle file selection
  handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const file = files[0];
      
      // Validate file type
      if (!file.type.startsWith('image/')) {
        this.setState({ error: '❌ Please select an image file' });
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

  // Upload image to KaizenImages library with KaizenID prefix
  uploadImage = async (file: File) => {
    this.setState({ uploading: true, message: 'Uploading...', error: '' });

    const libraryPath = this.getLibraryPath();
    const kaizenID = this.getKaizenID();
    
    // Create filename: [KaizenID]_[original filename]
    // Example: 9951_photo.jpg
    const fileExtension = file.name.substring(file.name.lastIndexOf('.'));
    const newFileName = `${kaizenID}_image${fileExtension}`;
    
    const encodedFileName = encodeURIComponent(newFileName);
    const uploadUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryPath}')/Files/add(url='${encodedFileName}',overwrite=true)`;

    try {
      const response = await this.props.context.spHttpClient.post(
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
    } catch (error) {
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

  // Delete image
  deleteImage = async () => {
    if (!window.confirm('Are you sure you want to delete this image?')) {
      return;
    }

    const libraryPath = this.getLibraryPath();
    const imageName = this.state.imageName;
    const encodedFileName = encodeURIComponent(imageName);
    
    const deleteUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryPath}')/Files('${encodedFileName}')`;

    try {
      this.setState({ uploading: true });
      
      await this.props.context.spHttpClient.post(
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
    } catch (error) {
      console.error('Delete error:', error);
      this.setState({ 
        error: '❌ Error deleting image',
        uploading: false 
      });
    }
  }

  render(): React.ReactElement<IImageUploadWebPartProps> {
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
```

---

## Step 6: ImageUploadWebPart.module.scss

**Path:** `src/webparts/imageUploadWebPart/components/ImageUploadWebPart.module.scss`

**Replace with:**

```scss
.imageUpload {
  padding: 20px;
  background: #ffffff;
  border-radius: 8px;
  max-width: 450px;
  border-left: 4px solid #0078d4;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);

  h3 {
    margin: 0 0 15px 0;
    color: #0078d4;
    font-size: 18px;
    font-weight: 600;
  }
}

.uploadButton {
  background-color: #0078d4;
  color: white;
  border: none;
  padding: 12px 20px;
  font-size: 14px;
  border-radius: 4px;
  cursor: pointer;
  font-weight: bold;
  width: 100%;
  transition: background-color 0.3s ease;
  margin-bottom: 15px;

  &:hover:not(:disabled) {
    background-color: #005a9e;
  }

  &:disabled {
    background-color: #ccc;
    cursor: not-allowed;
  }

  &:active:not(:disabled) {
    background-color: #004578;
  }
}

.message {
  margin-bottom: 15px;
  padding: 12px;
  border-radius: 4px;
  background-color: #d4edda;
  color: #155724;
  font-weight: bold;
  font-size: 13px;
  border-left: 3px solid #28a745;
}

.error {
  margin-bottom: 15px;
  padding: 12px;
  border-radius: 4px;
  background-color: #f8d7da;
  color: #721c24;
  font-weight: bold;
  font-size: 13px;
  border-left: 3px solid #f5222d;
}

.placeholder {
  color: #666;
  font-style: italic;
  padding: 20px;
  text-align: center;
  background-color: #f9f9f9;
  border-radius: 4px;
  border: 1px dashed #ddd;
  margin: 15px 0 0 0;
}

.imageContainer {
  margin-top: 20px;
  text-align: center;
  animation: slideIn 0.3s ease;
}

@keyframes slideIn {
  from {
    opacity: 0;
    transform: translateY(-10px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

.singleImage {
  width: 100%;
  max-height: 350px;
  object-fit: cover;
  border-radius: 4px;
  margin-bottom: 15px;
  border: 1px solid #e0e0e0;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.buttonGroup {
  display: flex;
  gap: 10px;
}

.deleteButton,
.replaceButton {
  flex: 1;
  padding: 10px;
  border: 1px solid #ddd;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  background-color: #f5f5f5;
  font-weight: 600;
  transition: all 0.3s ease;

  &:hover:not(:disabled) {
    background-color: #e8e8e8;
    border-color: #999;
  }

  &:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }
}

.deleteButton:hover:not(:disabled) {
  background-color: #ffebee;
  border-color: #f5222d;
  color: #f5222d;
}

.replaceButton:hover:not(:disabled) {
  background-color: #e3f2fd;
  border-color: #0078d4;
  color: #0078d4;
}
```

---

## Step 7: IImageUploadWebPartProps.ts

**Path:** `src/webparts/imageUploadWebPart/components/IImageUploadWebPartProps.ts`

```typescript
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IImageUploadWebPartProps {
  context: WebPartContext;
}
```

---

## Step 8: ImageUploadWebPart.ts

**Path:** `src/webparts/imageUploadWebPart/ImageUploadWebPart.ts`

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ImageUploadWebPartWebPartStrings';
import ImageUploadWebPart from './components/ImageUploadWebPart';
import { IImageUploadWebPartProps } from './components/IImageUploadWebPartProps';

export interface IImageUploadWebPartWebPartProps {
  description: string;
}

export default class ImageUploadWebPartWebPart extends BaseClientSideWebPart<IImageUploadWebPartWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IImageUploadWebPartProps> = React.createElement(
      ImageUploadWebPart,
      {
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
```

---

## Step 9: ImageUploadWebPart.manifest.json

**Path:** `src/webparts/imageUploadWebPart/ImageUploadWebPart.manifest.json`

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp-dev-declarative-web-parts/webpart.schema.json",
  "id": "3d9fc37f-5b1f-4e9f-8c2a-7e1b5f9d3a2c",
  "alias": "ImageUploadWebPart",
  "componentType": "WebPart",
  "displayName": "Image Upload",
  "description": "Upload and display a single image for Kaizen pages",
  "dataVersion": "1.0.0",
  "previewImageUrl": "",
  "searchablePropertyNames": [
    "Image",
    "Upload",
    "Kaizen"
  ],
  "version": "1.0.0",
  "manifestVersion": 2,
  "requiresCustomScript": false,
  "supportedHosts": [
    "SharePointWebPart"
  ],
  "properties": {}
}
```

---

# PART 4: BUILD & PACKAGE

## Step 10: Build
```bash
npm run build
```

## Step 11: Bundle
```bash
npm run bundle -- --ship
```

## Step 12: Package
```bash
npm run package-solution -- --ship
```

**Package location:** `sharepoint/solution/image-upload-webpart.sppkg`

---

# PART 5: DEPLOY

## Step 13: Go to App Catalog
1. https://[org]-admin.sharepoint.com
2. More features → App Catalog
3. Open "Apps for SharePoint"

## Step 14: Upload Package
1. Click + New → Files
2. Select `image-upload-webpart.sppkg`
3. Upload

## Step 15: Deploy
1. Click the uploaded file
2. Click "..." menu
3. Click "Deploy"
4. Check "Make this solution available to all sites"
5. Click Deploy
6. Wait 2-3 minutes

---

# PART 6: CREATE KaizenImages LIBRARY

## Step 16: Create Library

1. Go to your SharePoint site: `https://cranenxt.sharepoint.com/sites/NXT-CBS`
2. Click **+ New** → **Library**
3. Name it: **KaizenImages**
4. Click **Create**

This is where ALL images will be stored!

---

# PART 7: ADD WEB PART TO TEMPLATE

## Step 17: Edit Template
1. Go to SitePages library
2. Open your Kaizen template page
3. Click **Edit**

## Step 18: Add Web Part
1. Click **+ Add a new section**
2. Click **+** to add web part
3. Search: **"Image Upload"**
4. Add it

## Step 19: Save
1. Click **Save**

---

# PART 8: TEST

## Step 20: Test Upload
1. Create new page from template
2. Click "+ Add Image"
3. Select image
4. Should upload and display ✅

## Step 21: Verify Library
1. Go to **KaizenImages** library
2. Should see file: **9951_image.jpg** (or whatever KaizenID)
3. Each page's image named by its KaizenID ✅

---

# HOW IT WORKS

```
User opens: Kaizen9951
             ↓
         [+ Add Image]
             ↓
         Select image
             ↓
         Upload to KaizenImages library
         as: 9951_image.jpg
             ↓
         Display on page
             ↓
         KaizenImages Library now contains:
         ├── 9951_image.jpg ✅
         ├── 9950_image.jpg
         ├── 9949_image.jpg
         └── ...
```

---

# SUMMARY

✅ **Single centralized library** `KaizenImages`  
✅ **All images in one place**  
✅ **Auto-organized by KaizenID**  
✅ **Single image per page**  
✅ **Easy to manage**  
✅ **Works for all pages from template**  

**Clean and simple!** 🎉

