<?php 

    function Micelamnegrito_ubicacion_shortcode(){
        $ubicacion = get_field('ubicacion');

        ?>

    <div class= "ubicacion">
        <input type="hidden" id="lat" value="<?php echo $ubicacion['lat']; ?>" />
        <input type="hidden" id="lng" value="<?php echo $ubicacion['lng']; ?>" />
        <input type="hidden" id="direccion" value="<?php echo $ubicacion['address']; ?>" />

        <div id="mapa"></div>
        </div>

        <?php

    }

    add_shortcode('Micelamnegrito_ubicacion', 'Micelamnegrito_ubicacion_shortcode');

