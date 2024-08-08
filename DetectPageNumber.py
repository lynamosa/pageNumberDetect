import cv2
import os
import numpy as np

pos = open('movePage.js', 'w', encoding='utf8')

def draw_bounding_boxes(image_path):
    output_dir = 'pg2'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    filename = os.path.basename(image_path)
    
    image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)

    pageArea = 600 #px
    height, width = list(image.shape)

    # Tính toán vị trí và kích thước của hình 2x2cm ở phía dưới giữa
    x_start = max(0, width // 2 - int(pageArea/2))
    y_start = max(0, height - 300)
    x_end = min(width, width // 2 + int(pageArea/2))
    y_end = height

    # Trích xuất hình 2x2cm
    img = image[y_start:y_end, x_start:x_end]
    kernel = np.ones((3, 7), np.uint8)
    ret, thresh1 = cv2.threshold(img, 160, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY_INV)
    extracted_image = cv2.dilate(thresh1.copy(), kernel, iterations=3)
    # Tạo đường dẫn cho thư mục lưu trữ
    #output_path = os.path.join(output_dir, filename+'_thres.png')
    #cv2.imwrite(output_path, extracted_image)

    contours, _ = cv2.findContours(extracted_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Vẽ khung chữ nhật giới hạn cho từng vùng nội dung và thêm kích thước vào góc trên bên phải
    #for i, contour in enumerate(contours):
    #    x, y, w, h = cv2.boundingRect(contour)
    #    if h>40:
    #        cv2.rectangle(img, (x, y), (x + w, y + h), (0, 182, 0), 2)
    #        #cv2.putText(img, f'{w}x{h}', (x + w - 50, y + 20), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
    # Lọc ra các contours có chiều cao lớn hơn 40
    filtered_contours = [contour for contour in contours if cv2.boundingRect(contour)[3] > 40]
    
    if len(filtered_contours)>0:
        # Tìm contour có tọa độ x lớn nhất trong danh sách các contours được lọc
        max_x_contour = max(filtered_contours, key=lambda contour: cv2.boundingRect(contour)[1])

        # Tạo hình copy của extracted_image để vẽ khung chữ nhật lên đó
        extracted_image_with_rect = extracted_image.copy()

        # Vẽ khung chữ nhật cho contour có tọa độ x lớn nhất
        x, y, w, h = cv2.boundingRect(max_x_contour)
        cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 4)

        # Thêm kích thước vào góc trên bên phải
        cv2.putText(img, f'{w}x{h}', (x , y -10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
        pos.write(f'[{x+w/2},{y+h/2}],\n')
    else:
        pos.write(f'[0,0],\n')
        

    # Tạo đường dẫn cho thư mục lưu trữ
    output_path = os.path.join(output_dir, filename)
    cv2.imwrite(output_path, img)
    #print(output_path, end=', ')


if __name__ == "__main__":
    # os.chdir(r'D:\Sach\THanh Ca')
    # image_path = "img/SKM_C7592_Page_070_Image_0001.jpg"
    for x in os.listdir('img'):
        draw_bounding_boxes(f'img/{x}')
    print('FINISH!!!!!!')
