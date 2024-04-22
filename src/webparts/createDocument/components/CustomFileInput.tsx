import * as React from 'react';
import { useRef } from 'react';
import { IIconProps, PrimaryButton, } from '@fluentui/react';
import styles from './CreateDocument.module.scss';
interface CustomFileInputProps {
  onChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  key: number;
}
const CustomFileInput: React.FC<CustomFileInputProps> = ({ key, onChange }) => {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const chooseFileIcon: IIconProps = { iconName: "Attach" };
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    //const file = event.target.files?.[0] || undefined;
    onChange(event);
  };
  const handleClick = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };
  return (
    <div>
      <PrimaryButton  
       onClick={handleClick} 
       iconProps={chooseFileIcon} text={"Upload"} 
       className={`${styles.btn} ${styles.uploadDiv}`} />
      <input
        ref={fileInputRef}
        id="files"
        type="file"
        key={key}
        style={{ display: 'none' }}
        onChange={handleFileChange}
        multiple
        accept="
        audio/*,
        video/*,
        image/*,
        .doc, .docx, .pdf, .txt,
        .xls, .xlsx, application/vnd.ms-excel,
        .dwg, .dxf, .step, .stp, .igs, .iges, model/*"
      />
    </div>
  );
};
export default CustomFileInput;
