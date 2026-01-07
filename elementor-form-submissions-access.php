<?php
/****** Show Elementor Submissions to editors **/
if (!class_exists('ElementorFormSubmissionsAccess')) {

    class ElementorFormSubmissionsAccess
    {
        /** 
         * See if this user is just an editor (if they have edit_posts but not manage_options).
         * If they have manage_options, they can see the Form Submissions page anyway.
         * @return boolean
         */
        static function isJustEditor()
        {
            return current_user_can('edit_posts') && !current_user_can('manage_options');
        }

        /**
         * This is called around line 849 of wp-includes/rest-api/class-wp-rest-server.php by the ajax request which loads the data
         * into the form submissions view for Elementor (see the add_menu_page below). The ajax request checks the user has
         * the manage_options permission in modules/forms/submissions/data/controller.php within the handler's permission_callback.
         * This overrides that, and also for the call to modules/forms/submissions/data/forms-controller.php (which fills the
         * Forms dropdown on the submissions page). By changing the $route check below, you could open up more pages to editors.
         * @param array [endpoints=>handlers]
         * @return array [endpoints=>handlers]
         */
        static function filterRestEndpoints($endpoints)
        {
            if (mdp_user_can_access_menu('elementor-submissions') && !current_user_can('manage_options')) 
            {
                error_reporting(0); // Suppress PHP notices that could prevent JSON data from loading
                foreach ($endpoints as $route => $handlers) // For each endpoint
                    if (strpos($route, '/elementor/v1/form') === 0) // Check if it is one of the Elementor endpoints
                        foreach ($handlers as $num => $handler) // Loop through the handlers
                            if (is_array($handler) && isset($handler['permission_callback'])) // Check if this handler has a permission_callback
                                $endpoints[$route][$num]['permission_callback'] = function ($request) { return true; }; // Grant permission
            }
            return $endpoints;
        }

        /**
         * Add the submissions page to the admin menu on the left for editors only, as administrators
         * can already see it.
         */
        static function addOptionsPage()
        {
            if (!mdp_user_can_access_menu('elementor-submissions') || current_user_can('manage_options')) return;
            add_submenu_page(
                'statistiken',
                __('Anfragen', 'elementor-forms-statistics'),
                __('Anfragen', 'elementor-forms-statistics'),
                'edit_posts',
                'e-form-submissions',
                function () {
                    echo '<div id="e-form-submissions"></div>';
                }
            );
        }

        /**
         * Hook up the filter and action. I can't check if they are an editor here as the wp_user_can function
         * is not available yet.
         */
        static function hookIntoWordpress()
        {
            add_filter('rest_endpoints', array('ElementorFormSubmissionsAccess', 'filterRestEndpoints'), 1, 3);
            add_action('admin_menu', array('ElementorFormSubmissionsAccess', 'addOptionsPage'));
        }
    }

    ElementorFormSubmissionsAccess::hookIntoWordpress();
} // End of class wrapper
