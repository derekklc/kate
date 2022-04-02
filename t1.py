import pandas as pd

class mainCode:
    def compose(self):
        aa = pd.read_excel("Book1.xlsx", usecols=["Name", "Content", "note1", "note2"])
        # print(aa)
        print('------------------------')


        html = """
                    <div class="body_section_1">
                    <div class="body_title_1"> 
                """
        top_title = ""
        body_title_1 = ""
        body_img_1 = ""
        menu_1_buttons = []
        menu_1_body_items = []
        body_section_items = {}
        for i in range(0, len(aa.Name)):
            # print(aa.Name[i])
            if (aa.Name[i] == "top_title"):
                top_title = aa.Content[i]
            if (aa.Name[i] == "main_title"):
                body_title_1 = aa.Content[i]
            if (aa.Name[i] == "body_img_1"):
                body_img_1 = aa.Content[i]
            elif (aa.Name[i] == "menu_1_button"):
                menu_1_buttons.append(aa.Content[i])
            elif (aa.Name[i] == "menu_1_body_row_1"):
                menu_1_body_items.append({"label": aa.Content[i], "group": aa.note1[i]})
            elif (aa.Name[i] == "body_section_item_1"):
                group_name = "group_" + str(int(aa.note1[i]))
                if(not group_name in body_section_items):
                    body_section_items[group_name] = []
                body_section_items[group_name].append({"Content":aa.Content[i], "Type": aa.note2[i]})
                
        # print(f"body_section_items:: {body_section_items}")

        top_title_html = f"""<h3 class="bar">{top_title}</h3>"""

        body_img_1_html = f"""<div class="wplister-gallery">\n          <div class="gallery-thumbnail"><img src="{body_img_1}"></img></div></div> """

        body_title_1_html = f"""<div class="body_title_1">{body_title_1} </div>"""
        menu_1_button_html = ""
        # menu_1_button_html = f"""  <div class="body_menu_1">\n"""
        for i in range(0 ,len(menu_1_buttons)):
            checked = ""
            if(i == 0):
                checked = """ checked="checked" """
            menu_1_button_html += f"""  <input type="radio" id="item_{i+1}" name="item_1" {checked}></input>\n"""
            menu_1_button_html += f"""  <label for="item_{i+1}" class="body_menu_1_item">{menu_1_buttons[i]}</label>\n"""

        menu_1_body_html = f"""   <div class="body_content_1">\n """
        menu_1_body_groups = {}
        for i in range(0 ,len(menu_1_body_items)):
            current_group = int(menu_1_body_items[i]["group"])
            current_label = menu_1_body_items[i]["label"]

            # print(f"current_group::{current_group}")
            if (not current_group in menu_1_body_groups):
                if(not current_group == 1):
                    menu_1_body_html += """     </div>\n"""
                menu_1_body_html += """     <div class="menu_1_paragraph_1">\n"""

            menu_1_body_groups[current_group] = "group_exists"
            menu_1_body_html += f"""        <div class="menu_1_item_1">{current_label}</div>\n"""
            
        menu_1_body_html += "      </div>\n    </div>"

        body_html = """     <div id="feature_section_1">\n"""
        for group in body_section_items:
            for i in range(0, len(body_section_items[group])):
                tmp_type = body_section_items[group][i]["Type"]
                tmp_content = body_section_items[group][i]["Content"]
                if (tmp_type == "title"):
                    body_html += f"""            <div class="menu_1_title_1">{tmp_content}</div>\n"""
                elif (tmp_type == "img"):
                    body_html += f"""            <div class="menu_1_img_box_1">
                        <img class="menu_1_img_1" src="{tmp_content}"></img>
                    </div>\n"""
                elif (tmp_type == "description"):
                    body_html += f"""            <div class="menu_1_item_1">{tmp_content}</div>\n"""

        body_html += """    </div>\n"""

        wpl_wrapper_html = f"""
            <html>
                {self.getFixedPart1()}
                <div id="wpl_wrapper">
                    <div class="main_section">
                        {top_title_html}
                        <div id="wplister_gallery_container">
                            {body_img_1_html}
                            <div class="body_section_1">
                                {body_title_1_html}
                                <div class="body_menu_1">
                                    {menu_1_button_html}
                                    {menu_1_body_html}
                                </div>
                            </div>
                        </div>
                        {self.getFixedpart2()}
                        {body_html}
                        {self.getFixedPart3()}
                    </div>
                </div>
            </html>
        """


        print('------------------------')
        # print(f"body_title_1_html:: \n{top_title}\n")
        # print(f"body_title_1_html:: \n{body_title_1_html}\n")
        # print(f"menu_1_button_html:: \n{menu_1_button_html}\n")
        # print(f"menu_1_body_html:: \n{menu_1_body_html}\n")
        # print(f"body_html:: \n{body_html}\n")

        file = open('output.html', "w")
        # file.write(body_title_1_html)
        # file.write(menu_1_button_html)
        # file.write(menu_1_body_html)
        file.write(wpl_wrapper_html)

        file.close()

    def getFixedPart1(self):
        return """
        
        <style>
    .kb_paragraph_1 {
        display: flex;
        justify-content: flex-start;
        margin: 20px 0;
    }

    .kb_paragraph_1>img {
        width: 50vw;
        height: unset;
        min-width: 400px;
    }
    </style>
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <style type="text/css">
    /*
    Template: No Video product
    Description:
    Version: 1.2
    */
    /*
    Template: Has Video product
    Description:
    Version: 1.2
    */
    /* user colors */
    #wpl_store_header .hotline {
        color: #FFFFFF;
        background-color: #FFA500;
        display: block;
        float: right;
    }

    #wpl_wrapper .wpl_description h2 {
        color: #000000;
    }

    #wpl_wrapper h3.bar {
        color: #FFFFFF;
        background-color: #555555;
        background-image: -webkit-gradient(linear, left top, left bottom, from(#555555), to(#000000));
        background-image: -webkit-linear-gradient(top, #555555, #000000);
        background-image: -moz-linear-gradient(top, #555555, #000000);
        background-image: -ms-linear-gradient(top, #555555, #000000);
        background-image: -o-linear-gradient(top, #555555, #000000);
        background-image: linear-gradient(to bottom, #555555, #000000);
    }

    /* layout */
    #wpl_wrapper {
        width: 90%;
        margin-right: auto;
        margin-left: auto;
        margin-top: 30px;
        margin-bottom: 20px;
        font-size: 14px;
        font-family: Arial, Helvetica, sans-serif;
    }

    #wpl_wrapper h2 {
        font-size: 24px;
        line-height: 30px;
    }

    #wpl_wrapper h3.bar {
        margin-top: 30px;
        margin-bottom: 20px;
        padding-left: 10px;
        height: 30px;
        line-height: 30px;
        font-size: 16px;
        font-family: Helvetica, Arial, sans-serif;
        -webkit-border-radius: 5px;
        -moz-border-radius: 5px;
        border-radius: 5px;
    }

    #wpl_store_header {
        width: 100%;
        overflow: hidden;
    }

    #wpl_store_header .logo {
        float: left;
    }

    #wpl_store_header .hotline {
        height: 40px;
        font-size: 16px;
        line-height: 40px;
        font-weight: bold;
        padding-right: 20px;
        padding-left: 20px;
        margin-top: 7px;
        -webkit-border-radius: 5px;
        -moz-border-radius: 5px;
        border-radius: 5px;
        text-decoration: none;
        text-shadow: 2px 2px 4px #713803;
        filter: dropshadow(color=#713803, offx=0, offy=1);
    }

    #wpl_wrapper .clearfix {
        clear: both;
    }

    #wpl_list_images a {
        border: 1px solid #ccc;
        display: block;
        float: left;
        height: 75px;
        margin-bottom: 20px;
        margin-left: 10px;
        margin-right: 10px;
        padding: 12px;
        width: 75px;
    }

    #wpl_wrapper .alignleft {
        float: left;
    }

    #wpl_wrapper .alignright {
        float: right;
    }

    #wpl_wrapper .section p {
        padding-left: 10px;
        padding-right: 10px;
        /* fix for firefox: */
        line-height: 1.2em;
        margin: 1em 0;
    }

    #wpl_wrapper .section big {
        font-size: 18px;
        font-weight: bold;
    }

    .primary_image,
    .wpl_product_image {
        width: 100%;
    }

    #gallery a {
        text-decoration: none;
    }

    /* these only apply in preview to override default styles from wp-admin */
    body.wp-admin #wpl_wrapper ul {
        list-style: disc inside none;
    }

    body.wp-admin #wpl_wrapper ul li {
        padding-left: 2em;
    }

    /** Media queries **/
    @media (min-width: 320px) {
        #wpl_wrapper {
            width: 95%;
        }

        #wpl_store_header .hotline {
            float: left;
        }

        #gallery {
            width: 100%;
        }

        #wpl_list_images {
            display: none;
        }
    }

    @media (min-width: 481px) {
        #wpl_wrapper {
            width: 95%;
        }

        #wpl_store_header .hotline {
            float: right;
            display: inline-block;
        }

        #gallery {
            display: block;
            width: 48%;
            float: left;
        }

        #description {
            width: 48%;
            float: right;
        }

        #gallery #wpl_main_image {
            max-width: 220px;
            max-height: 225px;
        }

        #wpl_list_images {
            display: block;
        }

        #wpl_list_images a {
            width: 80px;
            height: 80px;
        }

        #wpl_list_images a img.wpl_thumb {
            width: 100%;
        }
    }

    @media (min-width: 641px) {
        #wpl_wrapper {
            width: 95%;
        }

        #wpl_store_header .hotline {
            float: right;
        }

        #gallery {
            display: block;
            width: 45%;
            float: left;
        }

        #description {
            width: 48%;
            float: right;
        }

        #gallery #wpl_main_image {
            width: 100%;
            height: 100%;
            max-height: 370px;
            max-width: 370px;
        }

        #wpl_list_images {
            display: block;
        }

        #gallery #wpl_list_images a {
            width: 80px;
            height: 80px;
        }

        #wpl_list_images a img.wpl_thumb {
            width: 80px;
        }
    }

    /* Style for the new product_gallery shortcode */
    .wplister-gallery {
        position: relative;
        width: 95%;
        margin: 2em auto auto
    }

    .wplister-gallery img {
        margin: 0 auto
    }

    .wplister-gallery {
        width: 48%;
        display: inline-block;
        padding-top: 480px;
        text-align: center;
    }

    .wplister-gallery .gallery-thumbnail {
        width: 100%;
        position: absolute;
        left: 0;
        top: 0;
        transition: all .5s;
        text-align: center;
        max-height: 570px;
        transition: all .5s;
        z-index: 22;
        height: 445px;
        background: #fff
    }

    .wplister-gallery .gallery-thumbnail img {
        max-width: 100%;
        max-height: 443px;
    }

    .wplister-gallery input[name='thumb_switch'] {
        display: none
    }

    .wplister-gallery label {
        margin-right: 11px;
        display: inline-block;
        cursor: pointer;
        transition: all .5s;
        opacity: 1;
        margin-bottom: 1em
    }

    .wplister-gallery span {
        display: table-cell;
        width: 90px;
        height: 90px;
        text-align: center;
        border: 1px solid #cdcdce;
        vertical-align: middle;
    }

    .wplister-gallery label img {
        max-width: 100%;
        width: auto;
        padding: 1px;
        max-height: 88px
    }

    .wplister-gallery input[name='thumb_switch']:checked+label {
        opacity: 1
    }

    .wplister-gallery input~.gallery-thumbnail {
        margin-bottom: 0
    }

    /* .wplister-gallery input[name='thumb_switch']~.gallery-thumbnail {
        opacity: 0;
        display: none
    } */

    .wplister-gallery input[name='thumb_switch']:checked+label+.gallery-thumbnail {
        opacity: 1;
        display: block;
        transform: scale(1)
    }

    .wplister-gallery #id1+label+.gallery-thumbnail {
        opacity: 1;
        display: block;
        transform: scale(1)
    }

    @media (max-width: 1210px) and (min-width: 1140px) {

        .wplister-gallery .gallery-thumbnail,
        .wplister-gallery .gallery-thumbnail img {
            max-height: 443px
        }
    }

    @media(min-width:992px) and (max-width:1140px) {

        .wplister-gallery #id1+label+.gallery-thumbnail:hover,
        .wplister-gallery input[name='thumb_switch']:checked+label+.gallery-thumbnail:hover {
            transform: scale(1.1);
            border: 1px solid #0084c9
        }
    }

    @media (max-width: 992px) {

        .wplister-gallery .gallery-thumbnail,
        .wplister-gallery .gallery-thumbnail img {
            max-height: 443px
        }
    }

    @media (max-width: 939px) {

        .wplister-gallery .gallery-thumbnail,
        .wplister-gallery .gallery-thumbnail img {
            max-height: 443px
        }
    }

    @media(width: 768px) {

        .wplister-gallery .gallery-thumbnail,
        .wplister-gallery .gallery-thumbnail img {
            max-height: 370px
        }

        .wplister-gallery label {
            margin-right: 14px
        }

        .wplister-gallery span {
            width: 81px;
            height: 81px
        }

        .wplister-gallery label img {
            max-height: 79px
        }
    }

    @media(max-width: 767px) {
        .wplister-gallery {
            width: 100%;
            position: relative
        }

        .wplister-gallery .gallery-thumbnail,
        .wplister-gallery .gallery-thumbnail img {
            max-height: 300px
        }

        .wplister-gallery label {
            margin-right: 12px
        }

        .wplister-gallery span {
            width: 66px;
            height: 66px
        }

        .wplister-gallery label img {
            max-height: 64px
        }
    }
    #wplister_gallery_container{
        display: flex;
    }
    #wplister_gallery_container > div {
        width: 50%;
    }
    .body_title_1 {
        font-size: 35px;
        padding: 15px 0;
    }
    </style>\n
        
        """

    def getFixedpart2(self):
        return """
        <style>
                input[name="item_1"]{
                display: none;
                }
                .menu_1_paragraph_1{
                display: none;
                padding: 5px 0;
                padding-left: 25px;
                border-left: 1px solid #999999;
                }
                input[name="item_1"]:checked + label {
                background: #999999;
                color: white;
                }
                input#item_1:checked ~ .body_content_1 > div:nth-child(1) {
                display: block;
                }
                input#item_2:checked ~ .body_content_1 > div:nth-child(2) {
                display: block;
                }
                input#item_3:checked ~ .body_content_1 > div:nth-child(3) {
                display: block;
                }
                input#item_4:checked ~ .body_content_1 > div:nth-child(4) {
                display: block;
                }
                .body_menu_1 {
                display: grid;
                grid-template-columns: 25% 25% 25% 25%;
                row-gap: 2px;
                }
                .body_title_1 {
                font-size: 35px;
                padding: 15px 0;
                padding-bottom: 25px;
                }

                .body_menu_1_item {
                width: calc(100% - 2px);
                text-align: center;
                border: 1px solid #999999;
                padding: 5px 0;
                font-size: calc(12px + 0.3vw);
                cursor: pointer;
                user-select: none;
                font-weight: normal;
                transition: 0.2s;
                }
                .body_menu_1_item:hover{
                background: #999999;
                color: white;
                }
                .menu_1_title_1 {
                line-height: 2;
                font-size: 20px;
                color: #393a4b;
                padding-top: 15px;
                margin-top: 25px;
                border-top: 1px solid grey;
                }
                .menu_1_item_1 {
                font-weight: normal;
                font-size: 16px;
                color: #4f5158;
                line-height: 2.5;
                text-align: center;
                }
                .menu_1_item_2 {
                font-weight: normal;
                padding: 5px 0;
                font-size: 16px;
                color: #4f5158;
                }

                .body_content_1 {
                grid-column-start: 1;
                grid-column-end: 5;
                margin-top: 3%;
                }
                .menu_1_img_1 {
                width: 100%;
                min-width: 200px;
                max-height: 800px;
                border-radius: 5px;
                margin: 10px 0;
                transition: 0.2s;
                }
                #feature_section_1 {
                text-align: center;
                padding: 25px 20vw 0 20vw;
                }

            </style>
        
        """

    def getFixedPart3(self):
        return """
            
        <div rwr="1" size="4" style="font-family:Arial">
        <h3 class="bar"
            style="color: rgb(255, 255, 255); background-color: rgb(85, 85, 85); background-image: linear-gradient(rgb(85, 85, 85), rgb(0, 0, 0)); margin-top: 30px; margin-bottom: 20px; padding-left: 10px; height: 30px; line-height: 30px; font-size: 16px; font-family: Helvetica, Arial, sans-serif; border-radius: 5px;">
            You May Like This</h3>
        <div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px;">
            <a href="https://www.ebay.com.au/itm/174604637245"
                title="Fantech PS5 Headset 3.5mm with Noise-Cancel Gaming Headphone for PC Switch Xbox" target="_top"
                style="float: left; display: block; width: 120px; height: 135px; margin-left: 0px; margin-right: 10px; border: 1px solid rgb(204, 204, 204); padding: 9px; color: rgb(0, 0, 153); text-decoration-line: none; border-radius: 5px; font-family: &quot;Times New Roman&quot;; font-size: medium; text-align: center;">
                <div class="thumb"><img src="https://i.ebayimg.com/images/g/CRYAAOSwGzpgB5R4/s-l1600.jpg"
                    alt="Fantech PS5 Headset 3.5mm with Noise-Cancel Gaming Headphone for PC Switch Xbox"
                    style="border: none; max-height: 90px; max-width: 120px; width: auto;"></div>
                <div class="title"
                    style="text-align: left; font-size: 12px; font-family: sans-serif; line-height: 15px; height: 45px; overflow: hidden; margin: 3px 0px;">
                    Fantech PS5 Headset 3.5mm with Noise-Cancel Gaming Headphone for PC Switch Xbox</div>
            </a>
            <a href="https://www.ebay.com.au/itm/174602020554"
                title="DarkFlash 550W Computer Power Supply 80 Plus Bronze Flat Cable ATX Desktop PSU" target="_top"
                style="float: left; display: block; width: 120px; height: 135px; margin-left: 0px; margin-right: 10px; border: 1px solid rgb(204, 204, 204); padding: 9px; color: rgb(0, 0, 153); text-decoration-line: none; border-radius: 5px; font-family: &quot;Times New Roman&quot;; font-size: medium; text-align: center;">
                <div class="thumb"><img src="https://i.ebayimg.com/images/g/LX8AAOSwzuNgBjDP/s-l1600.jpg"
                    alt="DarkFlash 550W Computer Power Supply 80 Plus Bronze Flat Cable ATX Desktop PSU"
                    style="border: none; max-height: 90px; max-width: 120px; width: auto;"></div>
                <div class="title"
                    style="text-align: left; font-size: 12px; font-family: sans-serif; line-height: 15px; height: 45px; overflow: hidden; margin: 3px 0px;">
                    DarkFlash 550W Computer Power Supply 80 Plus Bronze Flat Cable ATX Desktop PSU</div>
            </a>
            <a href="https://www.ebay.com.au/itm/174602021178"
                title="DarkFlash Gaming PC Case Tempered Glass ARGB Strip Micro-ATX Computer Tower Case" target="_top"
                style="float: left; display: block; width: 120px; height: 135px; margin-left: 0px; margin-right: 10px; border: 1px solid rgb(204, 204, 204); padding: 9px; color: rgb(0, 0, 153); text-decoration-line: none; border-radius: 5px; font-family: &quot;Times New Roman&quot;; font-size: medium; text-align: center;">
                <div class="thumb"><img src="https://i.ebayimg.com/images/g/vcAAAOSwKB1gBjU1/s-l1600.jpg"
                    alt="DarkFlash Gaming PC Case Tempered Glass ARGB Strip Micro-ATX Computer Tower Case"
                    style="border: none; max-height: 90px; max-width: 120px; width: auto;"></div>
                <div class="title"
                    style="text-align: left; font-size: 12px; font-family: sans-serif; line-height: 15px; height: 45px; overflow: hidden; margin: 3px 0px;">
                    DarkFlash Gaming PC Case Tempered Glass ARGB Strip Micro-ATX Computer Tower Case</div>
            </a>
            <a href="https://www.ebay.com.au/itm/174523786106"
                title="PC Desktop Gaming Mechanical Keyboard/Mouse/Mouse Pad RGB Backlit Computer Combo" target="_top"
                style="float: left; display: block; width: 120px; height: 135px; margin-left: 0px; margin-right: 10px; border: 1px solid rgb(204, 204, 204); padding: 9px; color: rgb(0, 0, 153); text-decoration-line: none; border-radius: 5px; font-family: &quot;Times New Roman&quot;; font-size: medium; text-align: center;">
                <div class="thumb"><img src="https://i.ebayimg.com/images/g/Wg0AAOSwsUJftg0N/s-l1600.jpg"
                    alt="PC Desktop Gaming Mechanical Keyboard/Mouse/Mouse Pad RGB Backlit Computer Combo"
                    style="border: none; max-height: 90px; max-width: 120px; width: auto;"></div>
                <div class="title"
                    style="text-align: left; font-size: 12px; font-family: sans-serif; line-height: 15px; height: 45px; overflow: hidden; margin: 3px 0px;">
                    PC Desktop Gaming Mechanical Keyboard/Mouse/Mouse Pad RGB Backlit Computer Combo</div>
            </a>
            <a href="https://www.ebay.com.au/itm/174520975488"
                title="PC Gaming Headset USB Wired Mic 7.1 Surround Sound RGB Light with stand bundle" target="_top"
                style="float: left; display: block; width: 120px; height: 135px; margin-left: 0px; margin-right: 10px; border: 1px solid rgb(204, 204, 204); padding: 9px; color: rgb(0, 0, 153); text-decoration-line: none; border-radius: 5px; font-family: &quot;Times New Roman&quot;; font-size: medium; text-align: center;">
                <div class="thumb"><img src="https://i.ebayimg.com/images/g/2JwAAOSwZ-9fsio2/s-l1600.jpg"
                    alt="PC Gaming Headset USB Wired Mic 7.1 Surround Sound RGB Light with stand bundle"
                    style="border: none; max-height: 90px; max-width: 120px; width: auto;"></div>
                <div class="title"
                    style="text-align: left; font-size: 12px; font-family: sans-serif; line-height: 15px; height: 45px; overflow: hidden; margin: 3px 0px;">
                    PC Gaming Headset USB Wired Mic 7.1 Surround Sound RGB Light with stand bundle</div>
            </a>
            <a href="https://www.ebay.com.au/itm/174672105355"
                title="DarkFlash ARGB 240mm Liquid Cooler CPU Water Cooling Intel 2066/1151 AMD AM4 AM3" target="_top"
                style="float: left; display: block; width: 120px; height: 135px; margin-left: 0px; margin-right: 10px; border: 1px solid rgb(204, 204, 204); padding: 9px; color: rgb(0, 0, 153); text-decoration-line: none; border-radius: 5px; font-family: &quot;Times New Roman&quot;; font-size: medium; text-align: center;">
                <div class="thumb"><img src="https://i.ebayimg.com/images/g/C7IAAOSw0QlgRggI/s-l1600.jpg"
                    alt="DarkFlash ARGB 240mm Liquid Cooler CPU Water Cooling Intel 2066/1151 AMD AM4 AM3"
                    style="border: none; max-height: 90px; max-width: 120px; width: auto;"></div>
                <div class="title"
                    style="text-align: left; font-size: 12px; font-family: sans-serif; line-height: 15px; height: 45px; overflow: hidden; margin: 3px 0px;">
                    DarkFlash ARGB 240mm Liquid Cooler CPU Water Cooling Intel 2066/1151 AMD AM4 AM3</div>
            </a>
            <a href="https://www.ebay.com.au/itm/175007428515"
                title="Fantech Gaming / Office Chair PU Leather 4D Armrest Recline Ergonomic (GC-283)" target="_top"
                style="float: left; display: block; width: 120px; height: 135px; margin-left: 0px; margin-right: 10px; border: 1px solid rgb(204, 204, 204); padding: 9px; color: rgb(0, 0, 153); text-decoration-line: none; border-radius: 5px; font-family: &quot;Times New Roman&quot;; font-size: medium; text-align: center;">
                <div class="thumb"><img src="https://i.ebayimg.com/images/g/aAIAAOSw2EBhg4fl/s-l1600.jpg"
                    alt="Fantech Gaming / Office Chair PU Leather 4D Armrest Recline Ergonomic (GC-283)"
                    style="border: none; max-height: 90px; max-width: 120px; width: auto;"></div>
                <div class="title"
                    style="text-align: left; font-size: 12px; font-family: sans-serif; line-height: 15px; height: 45px; overflow: hidden; margin: 3px 0px;">
                    Fantech Gaming / Office Chair PU Leather 4D Armrest Recline Ergonomic (GC-283)</div>
            </a>
            <a href="https://www.ebay.com.au/itm/175048048992"
                title="Fantech Home Office Chair Computer Breathable Mesh Ergonomic Computer Chair" target="_top"
                style="float: left; display: block; width: 120px; height: 135px; margin-left: 0px; margin-right: 10px; border: 1px solid rgb(204, 204, 204); padding: 9px; color: rgb(0, 0, 153); text-decoration-line: none; border-radius: 5px; font-family: &quot;Times New Roman&quot;; font-size: medium; text-align: center;">
                <div class="thumb"><img src="https://i.ebayimg.com/images/g/EJsAAOSwLi1hqatK/s-l1600.jpg"
                    alt="Fantech Home Office Chair Computer Breathable Mesh Ergonomic Computer Chair"
                    style="border: none; max-height: 90px; max-width: 120px; width: auto;"></div>
                <div class="title"
                    style="text-align: left; font-size: 12px; font-family: sans-serif; line-height: 15px; height: 45px; overflow: hidden; margin: 3px 0px;">
                    Fantech Home Office Chair Computer Breathable Mesh Ergonomic Computer Chair</div>
            </a>
        </div>
        <br><br><br><br>
        <p><span style="font-size: 14pt;">&nbsp;</span></p>
        <p><br></p>
        <table class="MsoTableGrid" border="0" cellspacing="0" cellpadding="0" width="699" style="width:900pt;background:#404040;mso-background-themecolor:text1;
        mso-background-themetint:191;border-collapse:collapse;border:none;mso-yfti-tbllook:
        1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:none;mso-border-insidev:
        none">
            <tbody>
                <tr style="mso-yfti-irow:0;mso-yfti-firstrow:yes;height:32.85pt">
                    <td width="350" valign="top" style="width:262.15pt;padding:0cm 5.4pt 0cm 5.4pt;
        height:32.85pt">
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal;mso-outline-level:
        3"><b style="font-size: 14pt;"><span style="font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1;
        letter-spacing:-.75pt"><br></span></b></p>
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal;mso-outline-level:
        3"><b style="font-size: 14pt;"><span style="font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1;
        letter-spacing:-.75pt">Sales</span></b></p>
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal;mso-outline-level:
        3">
                        <b><span style="font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1;
        letter-spacing:-.75pt">_________________________________________</span></b>
                        <b>
                            <span style="font-size:22.0pt;font-family:&quot;Arial&quot;,sans-serif;mso-fareast-font-family:
        &quot;Times New Roman&quot;;color:white;mso-themecolor:background1;letter-spacing:-.75pt">
                                <o:p></o:p>
                            </span>
                        </b>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l1 level1 lfo1;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1">
                            Please
                            do not bid or buy if you do not agree with our Terms of Sales, Payment
                            Options and Shipping Policies listed on this page.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l1 level1 lfo1;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1">
                            Entering
                            a bid or purchase a "Buy It Now" item signifies your agreement to
                            these terms.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l1 level1 lfo1;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1">
                            Please
                            contact us if you have any questions or problems before leaving a neutral or
                            negative feedback. We pride ourselves in our customer services. We also
                            believe that good communication can resolve most if not all problems.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l1 level1 lfo1;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1">
                            Shipment
                            will only be made after receiving CLEARED FULL payment. Items are despatched
                            Monday to Friday (excluding public holidays).
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l1 level1 lfo1;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1">
                            We
                            typically reply to emails within 24 hours (EXCLUDING weekends and public
                            holidays)
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l1 level1 lfo1;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1">
                            All
                            items come with 6 months limited warranty unless stated otherwise in the item
                            description. If your items arrived DOA (Dead On Arrival) or is defective (NOT
                            due to misuse), you must contact us within 7 days.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal">
                        <span style="color:white;mso-themecolor:background1">
                            <o:p>&nbsp;</o:p>
                        </span>
                    </p>
                    </td>
                    <td width="350" valign="top" style="width:262.15pt;padding:0cm 5.4pt 0cm 5.4pt;
        height:32.85pt">
                    <h3 style="margin-top:0cm;mso-outline-level:3"><span style="font-size:12.0pt;
        font-family:&quot;Arial&quot;,sans-serif;color:white;mso-themecolor:background1;
        letter-spacing:-.75pt"><br></span></h3>
                    <h3 style="margin-top:0cm;mso-outline-level:3">
                        <span style="font-size:12.0pt;
        font-family:&quot;Arial&quot;,sans-serif;color:white;mso-themecolor:background1;
        letter-spacing:-.75pt">
                            Payment
                            <o:p></o:p>
                        </span>
                    </h3>
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal;mso-outline-level:
        3">
                        <b><span style="font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1;
        letter-spacing:-.75pt">_________________________________________</span></b>
                        <b>
                            <span style="font-size:22.0pt;font-family:&quot;Arial&quot;,sans-serif;mso-fareast-font-family:
        &quot;Times New Roman&quot;;color:white;mso-themecolor:background1;letter-spacing:-.75pt">
                                <o:p></o:p>
                            </span>
                        </b>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l0 level1 lfo2;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            We accept Paypal and Bank Deposits for sales in Australia.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l0 level1 lfo2;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            Please include your ebay ID and item number as references if you pay using Bank Deposits. Failure to do so we will DELAY the order/shipping process significantly as we will NOT be able to match your payment to your purchases.                     <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l0 level1 lfo2;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            We do NOT accept cheques, postal order, money order or Cash On Delivery.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l0 level1 lfo2;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.0pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.0pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            Unpaid items will be reported to eBay and a strike will be placed on your account under eBay's Unpaid Item Process.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal">
                        <span style="color:white;mso-themecolor:background1">
                            <o:p>&nbsp;</o:p>
                        </span>
                    </p>
                    </td>
                </tr>
                <tr style="mso-yfti-irow:2;mso-yfti-lastrow:yes;height:31.0pt">
                    <td width="350" valign="top" style="width:262.15pt;padding:0cm 5.4pt 0cm 5.4pt;
        height:31.0pt">
                    <h3 style="margin-top:0cm;mso-outline-level:3">
                        <span style="font-size:10.0pt;
        font-family:&quot;Verdana&quot;,sans-serif;mso-bidi-font-family:Arial;color:white;
        mso-themecolor:background1;letter-spacing:-.75pt">
                            Feedback
                            <o:p></o:p>
                        </span>
                    </h3>
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal;mso-outline-level:
        3">
                        <b><span style="font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;
        mso-fareast-font-family:&quot;Times New Roman&quot;;color:white;mso-themecolor:background1;
        letter-spacing:-.75pt">_________________________________________</span></b>
                        <b>
                            <span style="font-size:22.0pt;font-family:&quot;Arial&quot;,sans-serif;mso-fareast-font-family:
        &quot;Times New Roman&quot;;color:white;mso-themecolor:background1;letter-spacing:-.75pt">
                                <o:p></o:p>
                            </span>
                        </b>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l2 level1 lfo5;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.5pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.5pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            We appreciate your positive feedback
                            and we will leave positive feedback immediately for you after receiving
                            yours.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l2 level1 lfo5;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.5pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.5pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            We would appreciate if you would
                            leave us positive feedback once you confirmed the item is received and in
                            good working order.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l2 level1 lfo5;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.5pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.5pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            If you are unhappy with our products
                            or services, please contact us before leaving a neutral or negative feedback.
                            We will try our utmost to resolve the issue to our mutual satisfaction.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l2 level1 lfo5;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.5pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.5pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            It would be much appreciated if you
                            can award us with "five stars" for "Detailed Seller
                            Ratings".
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-top:0cm;margin-right:3.75pt;margin-bottom:
        0cm;margin-left:26.25pt;text-indent:-18.0pt;line-height:normal;mso-list:l2 level1 lfo5;
        tab-stops:list 36.0pt">
                        <!--[if !supportLists]--><span style="font-size:10.0pt;
        mso-bidi-font-size:8.5pt;font-family:Symbol;mso-fareast-font-family:Symbol;
        mso-bidi-font-family:Symbol;color:white;mso-themecolor:background1"><span style="mso-list:Ignore">·<span
                                style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </span></span></span>
                        <!--[endif]-->
                        <span style="font-size:8.5pt;font-family:&quot;Arial&quot;,sans-serif;
        color:white;mso-themecolor:background1">
                            On the other hand, if you feel that
                            we do not achieve that level of performance, please feel free to contact us
                            as well.
                            <o:p></o:p>
                        </span>
                    </p>
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal">
                        <span style="color:white;mso-themecolor:background1">
                            <o:p>&nbsp;</o:p>
                        </span>
                    </p>
                    </td>
                    <td width="350" valign="top" style="width:262.15pt;padding:0cm 5.4pt 0cm 5.4pt;
        height:31.0pt">
                    <p class="MsoNormal" style="margin-bottom:0cm;line-height:normal">
                        <span style="color:white;mso-themecolor:background1">
                            <o:p>&nbsp;</o:p>
                        </span>
                    </p>
                    </td>
                </tr>
            </tbody>
        </table>
        <br>
        </div>\n
        """

x = mainCode()
x.compose()